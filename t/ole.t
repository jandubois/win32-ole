########################################################################
# If you rearrange the tests, please renumber:
# perl -i.bak -pe "++$t if !$t || s/^# \d+\./# $t./" ole.t
########################################################################

package Excel;
use Win32::OLE;

use strict qw(vars);
use vars qw($AUTOLOAD @ISA);
@ISA = qw(Win32::OLE);

sub AUTOLOAD {
  my $self = shift;
  $AUTOLOAD =~ s/.*::/SUPER::/;
  my $retval = $self->$AUTOLOAD(@_);
  return $retval if defined($retval) || $AUTOLOAD eq 'DESTROY';
  printf "# $AUTOLOAD returned OLE error 0x%08x\n", Win32::OLE->LastError();
  $::Fail = $::Test;
  return;
}

########################################################################

package main;
use strict;
use FileHandle;
use Win32::OLE qw(With);
use Win32::OLE::Const ('Microsoft Excel');
use Win32::OLE::Enum;
use vars qw($Test $Fail);

$^W = 1;
STDOUT->autoflush(1);
STDERR->autoflush(1);

open(ME,$0) or die $!;
my $TestCount = grep(/\+\+\$Test/,<ME>);
close(ME);

print STDERR "\n##### Ignore test failure if Excel is not installed #####\n";

my $File = Win32::GetCurrentDirectory . "\\test.xls";
unlink $File if -f $File;

$Test = 0;
print "1..$TestCount\n";

sub Quit {
  $_[0]->Win32::OLE::Quit;
  print "not " unless ++$Test == $TestCount;
  print "ok $TestCount\n";
}

# 1. Create a new Excel automation server
my $Excel = Excel->new('Excel.Application', \&Quit);
printf "# App object type is %s\n", Win32::OLE->QueryObjectType($Excel);
print "not " unless $Excel;
printf "ok %d\n", ++$Test;

# 2. Add a workbook (with default number of sheets)
my $Book = $Excel->Workbooks->Add or print "not ";
printf "# Book object type is %s\n", Win32::OLE->QueryObjectType($Book);
printf "ok %d\n", ++$Test;

# 3. Test if class is inherited by objects created through $Excel
print "not " unless UNIVERSAL::isa($Book,'Excel');
printf "ok %d\n", ++$Test;

# 4. Generate OLE error, should be "croaked" by Win32::OLE
eval { $Book->Xyzzy(223); };
chomp $@;
print "# Died with msg |$@|\n";
print "not " unless $@;
printf "ok %d\n", ++$Test;

# 5. Generate OLE error, should be trapped by Excel subclass
$Fail = -1;
{ local ($^W) = 0; $Book->Xyzzy(223); };
print "not " if $Fail != $Test;
printf "ok %d\n", ++$Test;

# 6. Get an object for 1st worksheet
my $Sheet = $Book->Worksheets(1) or print "not ";
printf "# Sheet object type is %s\n", Win32::OLE->QueryObjectType($Sheet);
printf "ok %d\n", ++$Test;

# 7. Test the "With" function
With($Sheet->PageSetup, Orientation => $xlLandscape, FirstPageNumber => 13);
my $Value = $Sheet->PageSetup->FirstPageNumber;
print "# FirstPageNumber is \"$Value\"\n";
print "not " unless $Value == 13;
printf "ok %d\n", ++$Test;

# 8. Test constant value: xlLandscape should be "2"
$Value = $Sheet->PageSetup->Orientation;
print "# Orientation is \"$Value\"\n";
print "not " unless $Value == 2;
printf "ok %d\n", ++$Test;

# 9. Call a method with a magical scalar as argument
my $Sheets = $Book->Worksheets;
my $Name = $Book->Worksheets($Sheets->{Count})->{Name};
print "# Name is \"$Name\"\n";
print "not " unless $Name;
printf "ok %d\n", ++$Test;

# 10. Set values of some cells and retrieve a value
$Sheet->{Name} = 'My Sheet #1';
foreach my $i (1..10) {
  $Sheet->Cells($i,$i)->{Value} = $i**2;
}
my $Cells = $Sheet->Cells(5,5);
printf "# Cells object type is %s\n", Win32::OLE->QueryObjectType($Cells);
$Value = $Cells->{Value};
print "# Value is \"$Value\"\n";
print "not " unless $Cells->{Value} == 25;
printf "ok %d\n", ++$Test;

# 11. Set a cell range from an array ref containing an IV, PV and NV
$Sheet->Range("A9:C9")->{Value} = [42, 'Perl', 3.1415];
$Value = $Sheet->Cells(9,2)->{Value};
print "# Value is \"$Value\"\n";
print "not " unless $Value eq 'Perl';
printf "ok %d\n", ++$Test;

# 12. Retrieve float value (esp. interesting in foreign locales)
$Value = $Sheet->Cells(9,3)->{Value};
print "# Value is \"$Value\"\n";
print "not " unless $Value == 3.1415;
printf "ok %d\n", ++$Test;

# 13. Set a cell formula and retrieve calculated value
my $xlCalculationAutomatic = -4105;
$Excel->{Calculation} = $xlCalculationAutomatic;
$Sheet->Cells(3,1)->{Formula} = '=PI()';
$Value = $Sheet->Cells(3,1)->{Value};
print "# Value is \"$Value\"\n";
print "not " unless abs($Value-3.141592) < 0.00001;
printf "ok %d\n", ++$Test;

# 14. Add single worksheet and check that worksheet count is incremented
my $Count = $Sheets->{Count};
$Book->Worksheets->Add;
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+1;
printf "ok %d\n", ++$Test;

# 15. Add 2 more sheets, optional arguments are omitted
$Count = $Sheets->{Count};
$Book->Worksheets->Add(undef,undef,2);
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+2;
printf "ok %d\n", ++$Test;

# 16. Add 3 more sheets before sheet 2 using a named argument
$Count = $Sheets->{Count};
$Book->Worksheets(2)->{Name} = 'XYZZY';
$Sheets->Add($Book->Worksheets(2), {Count => 3});
$Value = $Sheets->{Count};
print "# Count is \"$Count\" and Value is \"$Value\"\n";
print "not " unless $Value == $Count+3;
printf "ok %d\n", ++$Test;

# 17. Previous sheet 2 should now be sheet 5
$Value = $Book->Worksheets(5)->{Name};
print "# Value is \"$Value\"\n";
print "not " unless $Value eq 'XYZZY';
printf "ok %d\n", ++$Test;

# 18. Add 2 more sheets at the end using 2 named arguments
$Count = $Sheets->{Count};
# Following line doesn't work with Excel 7 (Seems like an Excel bug?)
# $Sheets->Add({Count => 2, After => $Book->Worksheets($Sheets->{Count})});
$Sheets->Add({Count => 2, After => $Book->Worksheets($Sheets->{Count}-1)});
print "not " unless $Sheets->{Count} == $Count+2;
printf "ok %d\n", ++$Test;

# 19. Number of objects in an enumeration must match its "Count" property
my @Sheets = Win32::OLE::Enum->All($Sheets);
printf "# \$Sheets->{Count} is %d\n", $Sheets->{Count};
printf "# scalar(\@Sheets) is %d\n", scalar(@Sheets);
foreach my $Sheet (@Sheets) {
  printf "# Sheet->{Name} is \"%s\"\n", $Sheet->{Name};
}
print "not " unless $Sheets->{Count} == @Sheets;
printf "ok %d\n", ++$Test;

# 20. Enumerate all application properties using the C<keys> function
my @Properties = keys %$Excel;
printf "# Number of Excel application properties: %d\n", scalar(@Properties);
$Value = grep /^(Parent|Xyzzy|Name)$/, @Properties;
print "# Value is \"$Value\"\n";
print "not " unless $Value == 2;
printf "ok %d\n", ++$Test;

# 21. Save workbook to file
print "not " unless $Book->SaveAs($File);
printf "ok %d\n", ++$Test;

# 22. Check if output file exists.
print "not " unless -f $File;
printf "ok %d\n", ++$Test;

# 23. Access the same file object through a moniker.
my $Obj = Win32::OLE->GetObject($File);
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

# 24. Terminate server instance ("ok $Test\n" printed by Excel destructor method)
exit;
