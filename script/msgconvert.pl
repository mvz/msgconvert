#!/usr/bin/perl -w
#
# msgconvert.pl:
#
# Convert .MSG files (made by Outlook (Express)) to multipart MIME messages.
#

use Config;
use Email::Outlook::Message;
use Email::Sender::Transport::Mbox;
use Getopt::Long;
use Pod::Usage;
use File::Basename;
use vars qw($VERSION);
$VERSION = "0.904";
my $XDG_OPEN_BIN = "/usr/bin/xdg-open";

# Setup command line processing.
my $verbose = '';
my $mboxfile = '';
my $destfolder = '';
my $open = '';
my $openwith = $XDG_OPEN_BIN;
my $help = '';	    # Print help message and exit.
GetOptions(
  'destfolder=s' => \$destfolder,
  'mbox=s' => \$mboxfile,
  'open' => \$open,
  'openwith' => \$openwith,
  'verbose' => \$verbose,
  'help|?' => \$help) or pod2usage(2);
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
    
    # If destfolder is used --open parameter is ignored
    if ($destfolder ne '') {
      $outfile = File::Spec->catfile($destfolder, "$basename.mime");
    } elsif ($open ne '' && $Config{osname} eq 'linux') { # Only used under linux
      $outfile = File::Spec->catfile("/tmp/", "$basename.mime");
    }
    open OUT, ">:utf8", $outfile
      or die "Can't open $outfile for writing: $!";
    binmode(OUT, ":utf8");
    print OUT $mail->as_string;
    close OUT;
    
    # File is opened
    if ($open ne '' && $Config{osname} eq 'linux') { # Only used under linux
      system($openwith, $outfile);
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
    --destfolder <folder>   folder where generated file will be placed
    --mbox <file>           deliver messages to mbox file <file>
    --open                  calls system's default program to open de converted file.
    --openwith              especify program to be used in order to open generated file.
    --verbose	            be verbose
    --help	                help message

=head1 OPTIONS

=over 8

=item B<--destfolder>

    Place generated .mime files into especified folder.

=item B<--mbox>

    Deliver to the given mbox file instead of creating individual .mime
    files.

=item B<--open>

    Calls /usr/bin/xdg-open (unless --openwith is used) once .mime file is 
    generated.

=item B<--openwith>

    Allows to define which program will be called after .mime file generation.
    
=item B<--verbose>

    Print information about skipped parts of the .msg file.

=item B<--help>

    Print a brief help message.

=head1 DESCRIPTION

This program will convert the messages contained in the Microsoft Outlook
files <file.msg>...  to message/rfc822 files with extension .mime.
Alternatively, if the --mbox option is present, all messages will be put in
the given mbox file.  This program will complain about unrecognized OLE
parts in the input files on stderr.

=head1 BUGS

The program will not check whether output files already exist. Also, if you
feed it "foo.MSG" and "foo.msg", you'll end up with one "foo.mime",
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
