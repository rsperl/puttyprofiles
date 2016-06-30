#! perl
#
# To use, put in a directory, such as c:erllib
#  Add this directory to your PERL5LIB variable:
#  1 - Right click on my computer and choose properties
#  2 - Click on "Advanced system settings"
#  3 - On the Advanced tab, click on "Environment Variables"
#  4 - Under "User variables for <username", choose "New"
#  5 - The variable name is PERL5LIB. The value is c:erllib or whatever
#      folder path you chose.
#  6 - Click ok, ok, and ok again.
#  Put PuTTYProfile.pm in the folder you created.
#  To test whether perl can find it, open a command prompt and type the
#  following:
#       perl -MPuTTYProfile -e 1
#  If nothing comes back, it works. If it can't find PuTTYProfile.pm, it
#  will complain.

###
### To set all profiles to look like the profile "My Default Profile":
###
#
# use PuTTYProfile;
# PuTTYProfile::set_all_profiles_to_default("My Default Profile");
#
#
###
### To create shortcuts for all profiles in ~\\PuTTY:
###
#
# Readonly my $DIR => $ENV{USERPROFILE} . '\\PuTTY';
# PuTTYProfile::create_shortcuts_for_all_sessions($DIR);
#
#
###
### To manipulate a single profile, you can do:
###
#
# my $putty = PuTTYProfile->new("profile_name");
# $putty->set_template_profile("My Default Profile");
# $putty->set_default_values();
#
#
###
### To create a shortcut to the profile:
###
#
# # optional call to set_path_to_putty -- only if PuTTY is
# # installed in a non-standard location
# $putty->set_path_to_putty('c:\ath\\to\utty.exe');
# $putty->create_shortcut($dir, $arg_string);
#
#
### Note that any setting in the array @EXCLUDED_KEYS will not be
### set. For example, you would never want to set the hostname.
#

package PuTTYProfile;

use strict;
use warnings;
use autodie;
use Win32::TieRegistry;
use Win32::OLE;
use Readonly;

Readonly my $PROFILES_ROOT =>
    'HKEY_CURRENT_USER\\Software\\SimonTatham\\PuTTY\\Sessions';
Readonly my @EXCLUDED_KEYS => qw(
    HostName
);
Readonly my $SINGLE_BACKSLASH => q(\\);
Readonly my $DOUBLE_BACKSLASH => q(\\\\);
Readonly my @PUTTY_LOCATIONS  => (
  "c:\\Program Files\\PuTTY\utty.exe",
  "c:\\Program Files (x86)\\PuTTY\utty.exe",
);

# Not an object method
sub get_all_profiles {
  my @sessions = keys %{ $Registry->{$PROFILES_ROOT} };
  return @sessions;
}

# Not an object method
sub dump_all_profiles {
  my ($save_dir) = @_;
  my @sessions = get_all_profiles();
  foreach my $session (@sessions) {
    my $key = $PROFILES_ROOT . "\\$session";
    $session =~ s/%20/ /g;
    my $putty = PuTTYProfile->new($session);
    $putty->dump($save_dir);
  }

}

# Not an object method
sub set_all_profiles_to_default {
  my ($template) = @_;
  my @sessions = &get_all_profiles();
  foreach my $session (@sessions) {
    $session =~ s/\\$//;
    my $session_decoded = $session;
    $session_decoded =~ s/%20/ /g;
    next if ( $session eq $template || $session_decoded eq $template );
    print "...$session_decoded\n";
    my $putty = PuTTYProfile->new($session);
    $putty->set_template_profile($template);
    $putty->set_default_values();
  }
  return;
}

# Not an object method
sub create_shortcuts_for_all_sessions {
  my ($dir) = @_;
  my @sessions = keys %{ $Registry->{$PROFILES_ROOT} };
  print "Creating shortcuts\n";
  foreach my $session (@sessions) {
    $session =~ s/\\$//;
    my $session_decoded = $session;
    $session_decoded =~ s/%20/ /g;
    next
        if ( $session eq 'Default Settings'
      || $session_decoded eq 'Default Settings' );
    my $putty = PuTTYProfile->new($session_decoded);
    print "...$session_decoded\n";
    $putty->create_shortcut($dir);
  }
  return;
}

sub new {
  my ( $class, $profile ) = @_;
  my $self = {};
  bless $self, $class;
  $self->{profile}         = $profile;
  $self->{profile_encoded} = $self->_encode_profile_name($profile);
  $self->{profiles_path}   = $PROFILES_ROOT;
  $self->{profile_path}
      = $self->{profiles_path} . $SINGLE_BACKSLASH . $self->{profile_encoded};
  return $self;
}

sub set_template_profile {
  my ( $self, $profile ) = @_;
  my $template_encoded = $profile;
  $self->{template_profile_encoded} = $self->_encode_profile_name($profile);
  $self->{template_profile}         = $profile;
  $self->{template_handle} = PuTTYProfile->new( $self->{template_profile} );
  return;
}

sub _encode_profile_name {
  my ( $self, $profile ) = @_;
  $profile =~ s/ /%20/g;
  return $profile;
}

sub get_keys {
  my ($self)      = @_;
  my $reg_profile = $self->{profile_path};
  my %keys        = %{ $Registry->{$reg_profile} };
  my @keys        = sort keys %keys;
  return @keys;
}

sub get_value {
  my ( $self, $key ) = @_;
  $key =~ s/^\\//;
  my @v = $Registry->{ $self->{profile_path} }->GetValue($key);
  return @v;
}

sub get_default_value {
  my ( $self, $key ) = @_;
  return $self->{template_handle}->get_value($key);
}

sub set_value {
  my ( $self, $key, $value, $type ) = @_;
  $key =~ s/^\\//;
  $Registry->{ $self->{profile_path} }->SetValue( $key, $value, $type );
  return;
}

sub set_default_value {
  my ( $self, $key ) = @_;
  my $check_key = $key;
  $check_key =~ s/^\\+//;
  return if grep {/^$check_key$/} @EXCLUDED_KEYS;
  my ( $default_value, $type ) = $self->get_default_value($key);
  $self->set_value( $key, $default_value, $type );
  return;
}

sub set_path_to_putty {
  my ( $self, $dir ) = @_;
  $self->{path_to_putty} = $dir;
}

sub _get_path_to_putty {
  my ($self) = @_;
  if ( !exists $self->{path_to_putty} ) {
    foreach my $path (@PUTTY_LOCATIONS) {
      $self->{path_to_putty} = $path if ( -f $path );
    }
  }
  if ( exists $self->{path_to_putty} ) {
    return $self->{path_to_putty};
  }
  else {
    die "Cannot find putty";
  }
}

sub create_shortcut {
  my ( $self, $dir, $argstring ) = @_;
  $argstring = '' unless ($argstring);
  my $lnk_path = $dir . '\\' . $self->{profile} . '.lnk';
  my $wsh      = new Win32::OLE 'WScript.Shell';
  my $shcut    = $wsh->CreateShortcut($lnk_path)
      or die "Can't create shortcut to $lnk_path: " . @!;
  $shcut->{TargetPath}  = $self->_get_path_to_putty();
  $shcut->{Arguments}   = '-load "' . $self->{profile} . '" ' . $argstring;
  $shcut->{Description} = 'Connect to ' . $self->{profile};
  $shcut->Save;
  undef $shcut;
  undef $wsh;
  return;
}

sub set_default_values {
  my ($self) = @_;
  my @keys = $self->get_keys;
  foreach my $key (@keys) {
    $self->set_default_value($key);
  }
  return;
}

sub dump {
  my ( $self, $save_dir ) = @_;
  my $profile = $self->{profile};
  $save_dir .= '\\' unless $save_dir =~ /\\$/;
  $profile =~ s/\\$//;
  my $filename = $save_dir . $profile . '.reg';
  my $key      = $self->{profile_path};
  my $cmd      = qq{regedit /E "$filename" "$key"};
  `$cmd`;
}

1;
