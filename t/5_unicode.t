########################################################################
# If you rearrange the tests, please renumber:
# perl -i.bak -pe "++$t if !$t || s/^# \d+\./# $t./" 5_unicode.t
########################################################################
#
# !!! These tests will not run unless "Unicode::String" is installed !!!
#
########################################################################

use strict;
use FileHandle;
use Win32::OLE::Variant;

$^W = 1;
STDOUT->autoflush(1);
STDERR->autoflush(1);

open(ME,$0) or die $!;
my $TestCount = grep(/\+\+\$Test/,<ME>);
close(ME);

eval { require Unicode::String };
if ($@) {
    print "# Unicode::String module not found.\n";
    print "1..0\n";
    exit 0;
}

my $Test = 0;
print "1..$TestCount\n";

# 1. Create a simple BSTR and convert to Unicode and back
my $v = Variant(VT_BSTR, '3,1415');
printf "# Type=%s Value=%s\n", $v->Type, $v->Value;
my $u = $v->Unicode;
print "not " unless $u->utf8 eq '3,1415';
printf "ok %d\n", ++$Test;

#print $u->hex, "\n";

