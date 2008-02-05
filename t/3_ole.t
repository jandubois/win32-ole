########################################################################
#
# Test Win32::OLE.pm compatibility module using MS Excel
#
########################################################################
# If you rearrange the tests, please renumber:
# perl -i.bak -pe "++$t if !$t || s/^# \d+\./# $t./" 3_ole.t
########################################################################

package Excel;
use Win32::OLE;

use strict qw(vars);
use vars qw($AUTOLOAD @ISA $Warn $LastError $CP $LCID);
@ISA = qw(Win32::OLE);

$CP   = $Win32::OLE::CP;
$LCID = $Win32::OLE::LCID;

sub AUTOLOAD {
  my $self = shift;
  $AUTOLOAD =~ s/.*::/SUPER::/;
  my $retval = $self->$AUTOLOAD(@_);
  return $retval if defined($retval) || $AUTOLOAD eq 'DESTROY';
  printf "# $AUTOLOAD returned OLE error 0x%08x\n", $LastError;
  $::Fail = $::Test;
  return;
}

########################################################################

package main;
use strict;
use FileHandle;
use Win32::OLE qw(CP_ACP CP_OEMCP in valof with OVERLOAD);
use Win32::OLE::Const ('Microsoft Excel');
use Win32::OLE::NLS qw(:DEFAULT :LANG :SUBLANG);
use Win32::OLE::Variant;
use vars qw($Test $Fail);

$^W = 1;

STDOUT->autoflush(1);
STDERR->autoflush(1);

open(ME,$0) or die $!;
my $TestCount = grep(/\+\+\$Test/,<ME>);
close(ME);

sub stringify {
    my $arg = shift;
    return "<undef>" unless defined $arg;
    if (ref $arg eq 'ARRAY') {
	my $res;
	foreach my $elem (@$arg) {
	    $res .= "," if defined $res;
	    $res .= stringify($elem);
	}
	return "[$res]";
    }
    return "$arg";
}

sub Quit {
  $_[0]->Win32::OLE::Quit;
  print "not " unless ++$Test == $TestCount;
  print "ok $TestCount\n";
}

# 1. Create a new Excel automation server
$Excel::Warn = 0;
my $Excel = Excel->new('Excel.Application', \&Quit);
$Excel::Warn = 2;
unless (defined $Excel) {
    print "# Excel.Application not installed!\n";
    my $Msg = Excel->LastError;
    chomp $Msg;
    $Msg =~ s/\n/\n\# /g;
    print "# $Msg\n";
    print "1..0\n";
    exit 0;
}

$Test = 0;
print "1..$TestCount\n";
my $File = Win32::GetCwd . "\\test.xls";
unlink $File if -f $File;

printf "# Excel is %s\n", overload::StrVal($Excel);
my $Type = Win32::OLE->QueryObjectType($Excel);
print "# App object type is $Type\n";
printf "ok %d\n", ++$Test;

# 2. Make sure the CreateObject function works too
my $Obj;
my $Value = Win32::OLE::CreateObject('Excel.Application', $Obj);
print "not " unless $Value && UNIVERSAL::isa($Obj, 'Win32::OLE');
printf "ok %d\n", ++$Test;
$Obj->Quit if defined $Obj;

# 3. Add a workbook (with default number of sheets)
my $Book = $Excel->Workbooks->Add;
$Type = Win32::OLE->QueryObjectType($Book);
print "# Book object type is $Type\n";
print "not " unless defined $Book;
printf "ok %d\n", ++$Test;

# 4. Test if class is inherited by objects created through $Excel
print "not " unless UNIVERSAL::isa($Book,'Excel');
printf "ok %d\n", ++$Test;

# 5. Generate OLE error, should be "croaked" by Win32::OLE
eval { local $Excel::Warn = 3; $Book->Xyzzy(223); };
my $Msg = $@;
chomp $Msg;
$Msg =~ s/\n/\n\# /g;
print "# Died with msg:\n# $Msg\n";
print "not " unless $@;
printf "ok %d\n", ++$Test;

# 6. Generate OLE error, should be trapped by Excel subclass
$Fail = -1;
{ local $Excel::Warn = 0; $Book->Xyzzy(223); };
printf "# Excel::LastError returns (num): %08x\n", Excel->LastError();
$Msg = Excel->LastError();
$Msg =~ s/\n/\n\# /g;
printf "# Excel::LastError returns (str):\n# $Msg\n";
Excel->LastError(0);
printf "# Excel::LastError returns (num): %08x\n", Excel->LastError();
printf "# Excel::LastError returns (str): %s\n", Excel->LastError();
print "not " if $Fail != $Test;
printf "ok %d\n", ++$Test;

# 7. Get an object for 1st worksheet
my $Sheet = $Book->Worksheets(1);
print "not " unless defined $Sheet;
$Type = Win32::OLE->QueryObjectType($Sheet);
print "# Sheet object type is $Type\n";
printf "ok %d\n", ++$Test;

# 8. Test the "With" function
with($Sheet->PageSetup, Orientation => xlLandscape, FirstPageNumber => 13);
$Value = $Sheet->PageSetup->FirstPageNumber;
print "# FirstPageNumber is \"$Value\"\n";
print "not " unless $Value == 13;
printf "ok %d\n", ++$Test;

# 9. Test constant value: xlLandscape should be "2"
$Value = $Sheet->PageSetup->Orientation;
print "# Orientation is \"$Value\"\n";
print "not " unless $Value == 2;
printf "ok %d\n", ++$Test;

# 10. Call a method with a magical scalar as argument
my $Sheets = $Book->Worksheets;
my $Name = $Book->Worksheets($Sheets->{Count})->{Name};
print "# Name is \"$Name\"\n";
print "not " unless $Name;
printf "ok %d\n", ++$Test;

# 11. Set values of some cells and retrieve a value
$Sheet->{Name} = 'My Sheet #1';
foreach my $i (1..10) {
  $Sheet->Cells($i,$i)->{Value} = $i**2;
}
my $Cell = $Sheet->Cells(5,5);
$Type = Win32::OLE->QueryObjectType($Cell);
print "# Cells object type is $Type\n";
$Value = $Cell->{Value};
print "# Value is \"$Value\"\n";
print "not " unless $Cell->{Value} == 25;
printf "ok %d\n", ++$Test;

# 12. Check if overloading conversion to number/string works
print "# Value is \"$Cell\"\n";
print "not " unless $Cell == 25;
printf "ok %d\n", ++$Test;

# 13. Test the valof function
my $RefOf = $Cell;
my $ValOf = valof $Cell;
$Cell->{Value} = 27;
print "not " unless $ValOf == 25 && $RefOf == 27;
printf "ok %d\n", ++$Test;

# 14. Set a cell range from an array ref containing an IV, PV and NV
$Sheet->Range("A8:C9")->{Value} = [[undef, 'Camel'],[42, 'Perl', 3.1415]];
$Value = $Sheet->Cells(9,2) . $Sheet->Cells(8,2);
print "# Value is \"$Value\"\n";
print "not " unless $Value eq 'PerlCamel';
printf "ok %d\n", ++$Test;

# 15. Retrieve float value (esp. interesting in foreign locales)
$Value = $Sheet->Cells(9,3)->{Value};
print "# Value is \"$Value\"\n";
print "not " unless $Value == 3.1415;
printf "ok %d\n", ++$Test;

# 16. Retrieve a 0 dimensional range; check array data structure
$Value = $Sheet->Range("B8")->{Value};
printf "# Values are: \"%s\"\n", stringify($Value);
print "not " if ref $Value;
printf "ok %d\n", ++$Test;

# 17. Retrieve a 1 dimensional row range; check array data structure
$Value = $Sheet->Range("B8:C8")->{Value};
printf "# Values are: \"%s\"\n", stringify($Value);
print "not " unless @$Value == 1 && ref $$Value[0];
printf "ok %d\n", ++$Test;

# 18. Retrieve a 1 dimensional column range; check array data structure
$Value = $Sheet->Range("B8:B9")->{Value};
printf "# Values are: \"%s\"\n", stringify($Value);
print "not " unless @$Value == 2 && ref $$Value[0] && ref $$Value[1];
printf "ok %d\n", ++$Test;

# 19. Retrieve a 2 dimensional range; check array data structure
$Value = $Sheet->Range("B8:C9")->{Value};
printf "# Values are: \"%s\"\n", stringify($Value);
print "not " unless @$Value == 2 && ref $$Value[0] && ref $$Value[1];
printf "ok %d\n", ++$Test;

# 20. Check contents of 2 dimensional array
$Value = $$Value[0][0] . $$Value[1][0] . $$Value[1][1];
print "# Value is \"$Value\"\n";
print "not " unless $Value eq 'CamelPerl3.1415';
printf "ok %d\n", ++$Test;

# 21. Set a cell formula and retrieve calculated value
$Sheet->Cells(3,1)->{Formula} = '=PI()';
$Value = $Sheet->Cells(3,1)->{Value};
print "# Value is \"$Value\"\n";
print "not " unless abs($Value-3.141592) < 0.00001;
printf "ok %d\n", ++$Test;

# 22. Add single worksheet and check that worksheet count is incremented
my $Count = $Sheets->{Count};
$Book->Worksheets->Add;
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+1;
printf "ok %d\n", ++$Test;

# 23. Add 2 more sheets, optional arguments are omitted
$Count = $Sheets->{Count};
$Book->Worksheets->Add(undef,undef,2);
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+2;
printf "ok %d\n", ++$Test;

# 24. Add 3 more sheets before sheet 2 using a named argument
$Count = $Sheets->{Count};
$Book->Worksheets(2)->{Name} = 'XYZZY';
$Sheets->Add($Book->Worksheets(2), {Count => 3});
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+3;
printf "ok %d\n", ++$Test;

# 25. Previous sheet 2 should now be sheet 5
$Value = $Book->Worksheets(5)->{Name};
print "# Value is \"$Value\"\n";
print "not " unless $Value eq 'XYZZY';
printf "ok %d\n", ++$Test;

# 26. Add 2 more sheets at the end using 2 named arguments
$Count = $Sheets->{Count};
# Following line doesn't work with Excel 7 (Seems like an Excel bug?)
# $Sheets->Add({Count => 2, After => $Book->Worksheets($Sheets->{Count})});
$Sheets->Add({Count => 2, After => $Book->Worksheets($Sheets->{Count}-1)});
print "not " unless $Sheets->{Count} == $Count+2;
printf "ok %d\n", ++$Test;

# 27. Number of objects in an enumeration must match its "Count" property
my @Sheets = in $Sheets;
printf "# \$Sheets->{Count} is %d\n", $Sheets->{Count};
printf "# scalar(\@Sheets) is %d\n", scalar(@Sheets);
foreach my $Sheet (@Sheets) {
    printf "# Sheet->{Name} is \"%s\"\n", $Sheet->{Name};
}
print "not " unless $Sheets->{Count} == @Sheets;
printf "ok %d\n", ++$Test;

# 28. Enumerate all application properties using the C<keys> function
my @Properties = keys %$Excel;
printf "# Number of Excel application properties: %d\n", scalar(@Properties);
$Value = grep /^(Parent|Xyzzy|Name)$/, @Properties;
print "# Value is \"$Value\"\n";
print "not " unless $Value == 2;
printf "ok %d\n", ++$Test;

# 29. Translate character from ANSI -> OEM
my ($Version) = $Excel->{Version} =~ /([0-9.]+)/;
print "# Excel version is $Version\n";

my $LCID = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_NEUTRAL));
$LCID = MAKELCID(MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL)) if $Version >= 8;
$Excel::LCID = $LCID;

$Cell = $Book->Worksheets('My Sheet #1')->Cells(1,5);
$Cell->{Formula} = '=CHAR(163)';
$Excel::CP = CP_ACP;
my $ANSI = valof $Cell;
$Excel::CP = CP_OEMCP;
my $OEM = valof $Cell;
print "# ANSI(cp1252) -> OEM(cp437/cp850): 163 -> 156\n";
print "# ANSI is \"$ANSI\" and OEM is \"$OEM\"\n";
print "not " unless ord($ANSI) == 163 && ord($OEM) == 156;
printf "ok %d\n", ++$Test;

# 30. Save workbook to file
print "not " unless $Book->SaveAs($File);
printf "ok %d\n", ++$Test;

# 31. Check if output file exists.
print "not " unless -f $File;
printf "ok %d\n", ++$Test;

# 32. Access the same file object through a moniker.
$Obj = Win32::OLE->GetObject($File);
for ($Count=0 ; $Count < 5 ; ++$Count) {
    my $Type = Win32::OLE->QueryObjectType($Obj);
    print "# Object type is \"$Type\"\n";
    last if $Type =~ /Workbook/;
    $Obj = $Obj->{Parent};
}
$Value = 2.7172;
eval { $Value = $Obj->Worksheets('My Sheet #1')->Range('A3')->{Value}; };
print "# Value is \"$Value\"\n";
print "not " unless abs($Value-3.141592) < 0.00001;
printf "ok %d\n", ++$Test;


# 33. Get return value as Win32::OLE::Variant object
$Cell = $Obj->Worksheets('My Sheet #1')->Range('B9');
my $Variant = Win32::OLE::Variant->new(VT_EMPTY, 0);
$Cell->Dispatch('Value', $Variant);
printf "# Variant is (%s,%s)\n", $Variant->Type, $Variant->Value;
print "not " unless $Variant->Type == VT_BSTR && $Variant->Value eq 'Perl';
printf "ok %d\n", ++$Test;

# 34. Use clsid string to start OLE server
undef $Value;
eval {
    use Win32::Registry;
    use vars qw($HKEY_CLASSES_ROOT);
    my ($HKey,$CLSID);
    $HKEY_CLASSES_ROOT->Create('Excel.Application\CLSID',$HKey);
    $HKey->QueryValue('', $CLSID);
    $Obj = Win32::OLE->new($CLSID);
    $Value = (Win32::OLE->QueryObjectType($Obj))[0];
};
print "# Object application is $Value\n";
print "not " unless $Value eq 'Excel';
printf "ok %d\n", ++$Test;

# 35. Use DCOM syntax to start server (on local machine though)
#     This might fail (on Win95/NT3.5 if DCOM support is not installed.
$Obj = Win32::OLE->new([Win32::NodeName, 'Excel.Application'], 'Quit');
$Value = (Win32::OLE->QueryObjectType($Obj))[0];
print "# Object application is $Value\n";
print "not " unless $Value eq 'Excel';
printf "ok %d\n", ++$Test;


# 36. Terminate server instance ("ok $Test\n" printed by Excel destructor)
exit;
