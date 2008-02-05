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

$Win32::OLE::Variant::CP = CP_ACP;
$Win32::OLE::Variant::LCID = $lcidEnglish;
printf "# LCID is %d\n", $Win32::OLE::Variant::LCID;
printf "# CP is %d\n", $Win32::OLE::Variant::CP;

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
$Win32::OLE::Variant::LCID = $lcidGerman;
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
$Win32::OLE::Variant::LCID = $lcidEnglish;
$v->ChangeType(VT_BSTR);
printf "# VT_BSTR Value in lcid=$lcidEnglish is \"%s\"\n", $v->As(VT_BSTR);
print "not " unless $v->Type == VT_BSTR && "$v" eq "3.1415";
printf "ok %d\n", ++$Test;

# 8. Try an invalid conversion and test LastError() method
$Win32::OLE::Variant::Warn = 0;
Win32::OLE::Variant->LastError(0);
my $Before = Win32::OLE::Variant->LastError;
$v = Variant(VT_BSTR, "Five");
$v->ChangeType(VT_I4);
printf "# Before: $Before After: %d\n", Win32::OLE::Variant->LastError;
print "not " unless $Before == 0 && Win32::OLE::Variant->LastError != 0;
printf "ok %d\n", ++$Test;
