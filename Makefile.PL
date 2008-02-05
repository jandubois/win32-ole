use Config;
use ExtUtils::MakeMaker qw(&WriteMakefile &neatvalue);

my @MM = 
(
 'NAME'	        => 'Win32::OLE',
 'VERSION_FROM' => 'OLE.pm',
 'PM'           => { 'OLE.pm'      => '$(INST_LIBDIR)/OLE.pm',
		     'Lite.pm'     => '$(INST_LIBDIR)/OLE/Lite.pm',
		     'Const.pm'    => '$(INST_LIBDIR)/OLE/Const.pm',
		     'Enum.pm'     => '$(INST_LIBDIR)/OLE/Enum.pm',
		     'NLS.pm'      => '$(INST_LIBDIR)/OLE/NLS.pm',
		     'TypeInfo.pm' => '$(INST_LIBDIR)/OLE/TypeInfo.pm',
		     'Variant.pm'  => '$(INST_LIBDIR)/OLE/Variant.pm',
		     # Compatibility layer: toplevel OLE.pm
		     'Compat.pm'   => '$(INST_LIBDIR)/../OLE.pm',
		   },
 'XS'	        => { 'OLE.xs' => 'OLE.cpp' },
 # Without the next line MakeMaker will not export the bootstrap function
 # (because we are using FUNCLIST too).
 'DL_FUNCS'     => { 'Win32::OLE' => [] },
 'dist'	        => { COMPRESS => 'gzip -9f', SUFFIX => 'gz' },
 'clean'	=> { FILES => 'test.xls' },
);

# Add additional settings for the creation of PPD files
if ($ExtUtils::MakeMaker::VERSION >= 5.43) {
    push @MM, 'ABSTRACT_FROM' => 'OLE.pm';
    push @MM, 'AUTHOR' => 'Jan Dubois (jan.dubois@ibm.net)';
    # Actually the OLE module was originally written by hip/ActiveWare.
    # It has been converted to standard Perl XS format by Gurusamy Sarathy.
    # Jan Dubois is currently the most active maintainer of this module.
    # Please see the OLE.pm file for complete copyright information.
}

# XXX Enable this when MakeMaker supports 'FUNCLIST'
if ($ExtUtils::MakeMaker::VERSION >= 5.4302) {
    push @MM, 'FUNCLIST' => [qw(SetSVFromVariant SetSVFromVariantEx SetVariantFromSV CreatePerlObject)];
}

my $FUNCLIST = [qw(SetSVFromVariant SetSVFromVariantEx SetVariantFromSV CreatePerlObject)];

WriteMakefile(@MM);

sub MY::xs_c {
    '
.xs.cpp:
	$(PERL) -I$(PERL_ARCHLIB) -I$(PERL_LIB) $(XSUBPP) $(XSPROTOARG) $(XSUBPPARGS) $*.xs >xstmp.c && $(MV) xstmp.c $*.cpp
';
}

sub MY::dlsyms {
    my($self,%attribs) = @_;

    my($funcs) = $attribs{DL_FUNCS} || $self->{DL_FUNCS} || {};
    my($vars)  = $attribs{DL_VARS} || $self->{DL_VARS} || [];
    my($imports)  = $attribs{IMPORTS} || $self->{IMPORTS} || {};
    my(@m);
    (my $boot = $self->{NAME}) =~ s/:/_/g;

    if (not $self->{SKIPHASH}{'dynamic'}) {
	push(@m,"
$self->{BASEEXT}.def: Makefile.PL
",
     q!	$(PERL) "-I$(PERL_ARCHLIB)" "-I$(PERL_LIB)" -MExtUtils::Mksymlists \\
     -e "Mksymlists('NAME' => '!, $self->{NAME},
     q!', 'DLBASE' => '!,$self->{DLBASE},
     q!', 'DL_FUNCS' => !,neatvalue($funcs),
     q!, 'FUNCLIST' => !,neatvalue($FUNCLIST),
     q!, 'IMPORTS' => !,neatvalue($imports),
     q!, 'DL_VARS' => !, neatvalue($vars), q!);"
!);
    }
    join('',@m);
}
