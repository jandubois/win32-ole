$self->{CC} = 'g++';
$self->{LIBS} = ['-lole32 -loleaut32 -luuid -lmsvcrt40'];
$self->{CCFLAGS} .= '-fvtable-thunks ' . $Config{ccflags};

# NOTE: These two functions are used for a typelib browser
#       that requires the ActiveState PerlScript wrapper.
sub MY::post_constants {}
sub MY::postamble {}
