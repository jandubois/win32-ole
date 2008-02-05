########################################################################
#
# Test of Win32::OLE::Variant
#
########################################################################
# If you rearrange the tests, please renumber:
# perl -i.bak -pe "++$t if !$t || s/^# \d+\./# $t./" 2_variant.t
########################################################################

use strict;
use FileHandle;
use Win32::OLE::NLS qw(:DEFAULT :LANG :SUBLANG);
use Win32::OLE::Variant qw(:DEFAULT CP_ACP);

$^W = 1;
STDOUT->autoflush(1);
STDERR->autoflush(1);

open(ME,$0) or die $!;
my $TestCount = grep(/\+\+\$Test/,<ME>);
close(ME);

my $Test = 0;
print "1..$TestCount\n";

my $lcidEnglish = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_NEUTRAL));
my $lcidGerman = MAKELCID(MAKELANGID(LANG_GERMAN, SUBLANG_NEUTRAL));

$Win32::OLE::CP = CP_ACP;
$Win32::OLE::LCID = $lcidEnglish;
printf "# LCID is %d\n", $Win32::OLE::LCID;
printf "# CP is %d\n", $Win32::OLE::CP;

# 1. Create a simple numeric variant
my $v = Variant(VT_R8, 3.1415);
print "not " unless UNIVERSAL::isa($v, 'Win32::OLE::Variant');
printf "ok %d\n", ++$Test;

# 2. Verify type and value of variant
printf "# Type is %d and Value is %f\n", $v->Type, $v->Value;
print "not " unless $v->Type == VT_R8 && $v->Value == 3.1415;
printf "ok %d\n", ++$Test;

# 3. Retrieve value as VT_BSTR value
printf "# As(VT_BSTR) is \"%s\"\n", $v->As(VT_BSTR);
print "not " unless $v->As(VT_BSTR) eq "3.1415";
printf "ok %d\n", ++$Test;

# 4. Change locale to "German" (uses ',' as decimal point)
$Win32::OLE::LCID = $lcidGerman;
printf "# As(VT_BSTR) in lcid=$lcidGerman is \"%s\"\n", $v->As(VT_BSTR);
print "not " unless $v->Value == 3.1415 && $v->As(VT_BSTR) eq "3,1415";
printf "ok %d\n", ++$Test;

# 5. Test overloaded conversion to string
printf "# String value is \"$v\"\n";
print "not " unless "$v" eq "3,1415";
printf "ok %d\n", ++$Test;

# 6. Test overloaded conversion to number
printf "# Numeric (0) value is %f\n", $v-3.1415;
print "not " unless abs($v-3.1415) < 0.00001;
printf "ok %d\n", ++$Test;

# 7. Change locale to "English" and convert VARIANT to VT_BSTR
$Win32::OLE::LCID = $lcidEnglish;
$v->ChangeType(VT_BSTR);
printf "# VT_BSTR Value in lcid=$lcidEnglish is \"%s\"\n", $v->As(VT_BSTR);
print "not " unless $v->Type == VT_BSTR && "$v" eq "3.1415";
printf "ok %d\n", ++$Test;

# 8. Try an invalid conversion and test LastError() method
$Win32::OLE::Warn = 0;
Win32::OLE->LastError(0);
my $Before = Win32::OLE->LastError;
$v = Variant(VT_BSTR, "Five");
$v->ChangeType(VT_I4);
printf "# Before: $Before After: %d\n", Win32::OLE->LastError;
print "not " unless $Before == 0 && Win32::OLE->LastError != 0;
printf "ok %d\n", ++$Test;

# 9. Special case: VT_UI1 with string argument implies VT_ARRAY
$v = Variant(VT_UI1, "Some string");
printf "# Type=%x String=\"%s\"\n", $v->Type, $v->Value;
print "not " unless $v->Type == VT_UI1|VT_ARRAY && $v->Value eq "Some string";
printf "ok %d\n", ++$Test;

# 10. Test assignment to specific type: float to I2
$v = Variant(VT_I2, 42);
printf "# Value (42) is %g\n", $v->Value;
$v->Put(3.1415);
printf "# Value (3.1415) is %g\n", $v->Value;
print "not " unless $v->Value == 3;
printf "ok %d\n", ++$Test;

# 11. Test assignment to specific type: large integer to I2
$v->Put(70_000);
printf "# Value (70_000) is %g\n", $v->Value;
print "not " unless $v->Value == 70_000-2**16;
printf "ok %d\n", ++$Test;

# 12. Test VT_BYREF using an alias pointing to the same VARIANT
my $t = Variant(VT_I4|VT_BYREF, 42);
$v = $t->Value;
printf "# Ref=%s Value=%s\n", ref($v), $v;
$v = $t->_Value; # NB: Undocumented and unsupported function for testing only!
printf "# Ref=%s Value=%s\n", ref($v), $v;
$t->Put(13);
printf "# Ref=%s Value=%s\n", ref($v), $v;
print "not " unless $v->Value == 13;
printf "ok %d\n", ++$Test;
undef $v;
undef $t;

# 13. Test various VT_UI1 manipulations
$v = Variant(VT_ARRAY|VT_UI1|VT_BYREF, 8);
$v->Put("1234567890");
$v->Put(1,'');
$v->Put(3,'ABC');
$v->Put(6,32);
printf "# String=\"%s\"\n", $v->Value;
print "not " unless $v->Value eq "1\0003A56 8";
printf "ok %d\n", ++$Test;

# 14. Assignment by string should be '\0' padded
$v->Put("ABCD");
printf "# String=\"%s\"\n", $v->Value;
print "not " unless $v->Value eq "ABCD"."\0" x 4;
printf "ok %d\n", ++$Test;

# 15. Test non-0 lower bound and Get() method
$v = Variant(VT_ARRAY|VT_UI1, [10,12]);
$v->Put("123");
printf "# String=\"%s\", Get(11)=%d\n", $v->Value, $v->Get(11);
print "not " unless $v->Get(11) == ord('2');
printf "ok %d\n", ++$Test;

# 16. Test multidimensional array
$v = Variant(VT_ARRAY|VT_VARIANT, 3, [1,2]);
my @dim = $v->Dim;
print "#";
foreach (@dim) { printf " [%s]", join(',', @$_); }
print "\n";
print "not " unless $dim[0][0] == 0 && $dim[0][1] == 2 && $dim[1][0] == 1 && $dim[1][1] == 2;
printf "ok %d\n", ++$Test;

# 17. Assignment to VT_VARIANT array
$v->Put(0, 1, "Perl");
$v->Put(1, 2, 3.1415);
printf "# String=\"%s\" Number=%s\n", $v->Get(0,1), $v->Get(1,2);
print "not " unless $v->Get(0,1) eq 'Perl' && $v->Get(1,2) == 3.1415;
printf "ok %d\n", ++$Test;
