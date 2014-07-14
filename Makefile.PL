use 5.006;
use ExtUtils::MakeMaker;
WriteMakefile(
  'NAME'	    => 'msgconvert',
  'AUTHOR'	    => 'Matijs van Zuijlen',
  'VERSION_FROM'    => 'tools/msgconvert.pl',
  'EXE_FILES'	    => [ 'tools/msgconvert.pl', 'tools/oledump.pl' ],
  'PREREQ_PM'	=> {
    Email::Outlook::Message => 0.901,
    OLE::Storage_Lite => 0.14,
    Email::Sender => 1.3,
    Getopt::Long => 0,
    Pod::Usage => 0,
    File::Basename => 0,
  },
);

1;
