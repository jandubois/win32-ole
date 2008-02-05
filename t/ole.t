
#
# This is just a wrapper class that provides a DESTROY method
# that will close active connections gracefully
#
package Excel;
#use blib;
use Win32::OLE;

print STDERR "\n##### Ignore test failure if Excel is not installed #####\n";
print STDERR "\n#####       Click on 'Yes' for all the dialogs      #####\n";

sub new {
    my $s = {};
    if ($s->{Ex} = Win32::OLE->new('Excel.Application')) {
	return bless $s, shift;
    }
    return undef;
}

sub DESTROY {
    my $s = shift;
    if (exists $s->{Ex}) {
        print "# closing connection\n";
	print "ok 5\n";
	$s->{Ex}->Quit;
	return undef;
    }
}
sub AUTOLOAD {
    my $s = shift;
    $AUTOLOAD =~ s/^.*:://;
    $s->{Ex}->$AUTOLOAD(@_);
}

package main;

$ex = Excel->new or die "# Excel unavailable, skipping tests\n1..0\n";

print "1..5\n";

print "not " unless $ex->Workbooks->Add;
print "ok 1\n";
# dying after establishing a connection causes the server to hang around
# (process leak), which is why we do a wrapper class with its own DESTROY.
#die "stopped\n";

my $exs = $ex->Workbooks(1)->Worksheets('Sheet1') or print "not ";
print "ok 2\n";
for (1..10) {
    $exs->Cells($_,$_)->{Value} = $_**2;
    #$ex->Workbooks(1)->Worksheets('Sheet1')->Cells($_,$_)->{Value} = $_**2;
}
print "not " unless $exs;
print "ok 3\n";
print "not " unless $ex->Save();
print "ok 4\n";

