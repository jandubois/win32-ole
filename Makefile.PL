use ExtUtils::MakeMaker;
WriteMakefile(
    'NAME'	   => 'Win32::OLE',
    'VERSION_FROM' => 'OLE.pm', # finds $VERSION
    'PREREQ_PM'    => { 'Win32::Registry' => 0 },
    'PM'           => { 'OLE.pm'     => '$(INST_LIBDIR)/OLE.pm',
			'Const.pm'   => '$(INST_LIBDIR)/OLE/Const.pm',
			'Enum.pm'    => '$(INST_LIBDIR)/OLE/Enum.pm',
			'Variant.pm' => '$(INST_LIBDIR)/OLE/Variant.pm',
		      },
    'XS'	   => { 'OLE.xs' => 'OLE.cpp' },
    'dist'	   => { COMPRESS => 'gzip -9f', SUFFIX => 'gz' },
    'clean'	   => { FILES => 'test.xls' },
);

sub MY::xs_c {
    '
.xs.cpp:
	$(PERL) -I$(PERL_ARCHLIB) -I$(PERL_LIB) $(XSUBPP) $(XSPROTOARG) $(XSUBPPARGS) $*.xs >xstmp.c && $(MV) xstmp.c $*.cpp
';
}
