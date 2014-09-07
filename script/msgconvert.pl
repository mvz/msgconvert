#!/usr/bin/perl -w
#
# msgconvert.pl:
#
# Convert .MSG files (made by Outlook (Express)) to multipart MIME messages.
#

use Config;
use Email::Outlook::Message;
use Email::Sender::Transport::Mbox;
use Getopt::Long qw(:config no_auto_abbrev bundling);
use Pod::Usage;
use File::Basename;
use File::Temp qw/ tempfile /;
use vars qw($VERSION);
$VERSION = "0.905";

my $XDG_OPEN_BIN = '';
# Find standard xdg-open on UNIX-like systems
if ($Config{osname} ne 'darwin|MSWin32|cygwin'){
    my $xdgOpenPath = `which xdg-open`;
    chomp($xdgOpenPath);
	$XDG_OPEN_BIN = $xdgOpenPath;
}


# Setup command line processing.
my $verbose = '';
my $mboxfile = '';
my $destfolder = '';
my $open = '';
my $openwith = '';
my $help = '';	    # Print help message and exit.
GetOptions(
  'd|destfolder=s' => \$destfolder,
  'm|mbox=s' => \$mboxfile,
  'o|open' => \$open,
  'w|openwith=s' => \$openwith,
  'v|verbose' => \$verbose,
  'h|help|?' => \$help) or pod2usage(2);
pod2usage(1) if $help;

# Check file names
defined $ARGV[0] or pod2usage(2);

my $using_mbox = $mboxfile ne '';
my $transport;

if ($using_mbox) {
  $transport = Email::Sender::Transport::Mbox->new({ filename => $mboxfile });
}

foreach my $file (@ARGV) {
  my $msg = new Email::Outlook::Message($file, $verbose);
  my $mail = $msg->to_email_mime;
  if ($using_mbox) {
    $transport->send($mail, { from => $mail->header('From') || '' });
  } else {
    my $basename = basename($file, qr/\.msg/i);
    my $outfile = "$basename.eml";
    my $basepath = '';

    # If no --destfolder is specified and --open or --openwith is used, 
    # OS dependant tempFile will be used.
    if (($open ne '' || $openwith ne '') && $destfolder eq '') {
        ($tf, $outfile)=tempfile();
        close $tf;
        $outfile="$outfile.eml";
    } else { 
        if ($destfolder ne '') {
            $basepath=$destfolder;
        } else {
            $basepath=".";
        }
        $outfile = File::Spec->catfile($basepath, "$basename.eml");
    }
    
    open OUT, ">:utf8", $outfile
      or die "Can't open $outfile for writing: $!";
    binmode(OUT, ":utf8");
    print OUT $mail->as_string;
    close OUT;

    # File is opened
    my $customFileOpen = $XDG_OPEN_BIN;
    # First we need to find the correct way of launching generated mail on each
    # OS
    if($Config{osname} eq 'darwin' && $openwith eq '') { # MacOS
        $customFileOpen='open';
    } elsif ($Config{osname} eq 'MSWin32|cygwin' && $openwith eq ''){ #Win
        $customFileOpen='start';
    } elsif ($openwith ne '') {
        $customFileOpen=$openwith;
    }
    # Launching file if --open or --openwith is have been used
    if ($open ne '' || $openwith ne ''){
        system($customFileOpen, $outfile);
    }
  }
}

#
# Usage info follows.
#
__END__

=head1 NAME

msgconvert.pl - Convert Outlook .msg files to mbox format

=head1 SYNOPSIS

msgconvert.pl [options] <file.msg>...

  Options:
    -d, --destfolder=FOLDER     folder where generated file will be placed
    -m, --mbox=FILE             deliver messages to mbox file <file>
    -o, --open                  calls system's default program to open de converted file.
    -w, --openwith=PROGRAM_PATH especify program to be used in order to open generated file.
    -v, --verbose               be verbose
    -h, --help                  help message

=head1 OPTIONS

=over 8

=item B<-d, --destfolder>

    Place generated .eml files into especified folder.

=item B<-m, --mbox>

    Deliver to the given mbox file instead of creating individual .eml
    files.

=item B<-o, --open>

    Calls open,start or xdg-open (depending on OS) unless --openwith is used 
    once .eml file is generated.

=item B<-w, --openwith>

    Allows to define which program will be called after .eml file generation.
    
=item B<-v, --verbose>

    Print information about skipped parts of the .msg file.

=item B<-h, --help>

    Print a brief help message.

=head1 DESCRIPTION

This program will convert the messages contained in the Microsoft Outlook
files <file.msg>...  to message/rfc822 files with extension .eml.
Alternatively, if the --mbox option is present, all messages will be put in
the given mbox file.  This program will complain about unrecognized OLE
parts in the input files on stderr.

--open and --openwith parameters allows to integrate it into desktop
environments and use it to open generated file directly with your favourite 
application.

=head1 BUGS

The program will not check whether output files already exist. Also, if you
feed it "foo.MSG" and "foo.msg", you'll end up with one "foo.eml",
containing one of the messages.

Not all data that's in the .MSG file is converted. There simply are some
parts whose meaning escapes me. One of these must contain the date the
message was sent, for example. Formatting of text messages will also be
lost. YMMV.

=head1 AUTHOR

Matijs van Zuijlen, C<matijs@matijs.net>

=head1 COPYRIGHT AND LICENSE

Copyright 2002, 2004, 2006, 2007 by Matijs van Zuijlen

This program is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
