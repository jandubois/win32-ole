use ExtUtils::MakeMaker;
WriteMakefile(
    'NAME'	=> 'Win32::OLE',
    'VERSION_FROM' => 'OLE.pm', # finds $VERSION
    'XS'	=> { 'OLE.xs' => 'OLE.cpp' },
    'dist'	=> { COMPRESS => 'gzip -9f', SUFFIX => 'gz' },
);

sub MY::xs_c {
    '
.xs.cpp:
	$(PERL) -I$(PERL_ARCHLIB) -I$(PERL_LIB) $(XSUBPP) $(XSPROTOARG) $(XSUBPPARGS) $*.xs >xstmp.c && $(MV) xstmp.c $*.cpp
';
}
