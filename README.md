# MSGConvert: A .MSG to mbox convertor

**Note: `msgconvert` is now part of Email::Outlook::Message. Please visit the
[Email::Outlook::Message repository on
GitHub](https://github.com/mvz/email-outlook-message-perl) for the latest
source code.**

## Usage

To use it, run:

    perl -w msgconvert.pl YourMessage.msg

This will produce a file YourMessage.mime containing the message in RFC822
format. The program will complain about unrecognized OLE parts and other
problems on stderr. If you supply the option `--verbose`, it will also tell you
what OLE parts it knows about but doesn't use, and what I think they are. The
option `--help` will make it print some usage information.

You can also let MSGConvert deliver all .MSG files in one mbox file using
`--mbox`, like so (assuming you made `msgconvert.pl` executable):

    msgconvert.pl --mbox some-mbox-file *.msg

## Installing

The following instructions should work generally on any Unix-like system with a
reasonably new Perl installed.

* Download the script and put it in a convenient location.
* Install the necessary Perl modules by executing the following:

    cpan -i Email::Sender Email::Outlook::Message

  You can run this as root if you would like to install these modules system-wide.

On Debian and Ubuntu, try the following:

* Download the script and put it in a convenient location.
* Install Email::Outlook::Message and Email::Sender by executing the following:

    sudo apt-get install libemail-outlook-message-perl libemail-localdelivery-perl

# Development

For the latest source code, go to
[MSGConvert on GitHub](https://github.com/mvz/msgconvert) and
[Email::Outlook::Message on GitHub](https://github.com/mvz/email-outlook-message-perl).

# Known Bugs/Issues

Not all data that's in the .MSG file is converted. There simply are some parts
whose meaning escapes me. However, most things are converted correctly by now,
including plain text, HTML and RTF-formatted message bodies.

Attachments with Apple-style resource forks, as well as PGP-signed email is
known not to be converted properly.
