use Config;
use ExtUtils::MakeMaker qw(&WriteMakefile &neatvalue);

my @MM =
(
 'NAME'	        => 'Win32::OLE',
 'VERSION_FROM' => 'lib/Win32/OLE.pm',
 'XS'	        => { 'OLE.xs' => 'OLE.cpp' },
 # Without the next line MakeMaker will not export the bootstrap function
 # (because we are using FUNCLIST too).
 'DL_FUNCS'     => { 'Win32::OLE' => [] },
 'dist'	        => { COMPRESS => 'gzip -9f', SUFFIX => 'gz' },
 'clean'	=> { FILES => 'test.xls browser_to_blib' },
);

# Add additional settings for the creation of PPD files
if ($ExtUtils::MakeMaker::VERSION >= 5.43) {
    push @MM, 'ABSTRACT_FROM' => 'lib/Win32/OLE.pm';
    push @MM, 'AUTHOR' => 'Jan Dubois (jand@activestate.com)';
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

# Install typelib browser for ActivePerl help
sub MY::post_constants {
    my @files = qw(Browser.dhtml Browser.html Class.png Const.png Default.png
                   Enum.png Event.png Function.png Global.png Help.png
                   Library.png Module.png Property.png Type.png);
    my $m = '
INST_BROWSER = blib\html\lib\site\Win32\OLE
BROWSER_FILES =';
    $m .= "\t\\\n\tbrowser\\$_" foreach @files;
    $m .= '

BROWSER_TO_BLIB =';
    $m .= "\t\\\n\tbrowser\\$_ \$(INST_BROWSER)\\$_" foreach @files;
    return $m."\n\n";
}

sub MY::postamble {
    return q{
all :: browser_to_blib

browser_to_blib : $(BROWSER_FILES)
	@$(MKPATH) $(INST_BROWSER)
	@$(PERL) "-I$(INST_ARCHLIB)" "-I$(INST_LIB)" \
	"-I$(PERL_ARCHLIB)" "-I$(PERL_LIB)" -MExtUtils::Install \
        -e "pm_to_blib(qw[ <<pmfiles.dat ],'$(INST_LIB)\auto')"
$(BROWSER_TO_BLIB)
<<
	@$(TOUCH) $@

};
}

