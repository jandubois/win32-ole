#
# Documentation at the __END__
#

package Win32::OLE;

$VERSION = '0.03';
#use Carp;
use Exporter;
use DynaLoader;
@ISA = qw( Exporter DynaLoader );

@EXPORT = qw(
		Variant
		VT_UI1
		VT_I2
		VT_I4
		VT_R4
		VT_R8
		VT_DATE
		VT_BSTR
		VT_CY
		VT_BOOL
	    );
@EXPORT_OK = qw(
		VT_EMPTY
		VT_NULL
		VT_DISPATCH
		VT_ERROR
		VT_VARIANT
		VT_UNKNOWN
		VT_UI2
		VT_UI4
		VT_I8
		VT_UI8
		VT_INT
		VT_UINT
		VT_VOID
		VT_HRESULT
		VT_PTR
		VT_SAFEARRAY
		VT_CARRAY
		VT_USERDEFINED
		VT_LPSTR
		VT_LPWSTR
		VT_FILETIME
		VT_BLOB
		VT_STREAM
		VT_STORAGE
		VT_STREAMED_OBJECT
		VT_STORED_OBJECT
		VT_BLOB_OBJECT
		VT_CF
		VT_CLSID
		TKIND_ENUM
		TKIND_RECORD
		TKIND_MODULE
		TKIND_INTERFACE
		TKIND_DISPATCH
		TKIND_COCLASS
		TKIND_ALIAS
		TKIND_UNION
		TKIND_MAX
	       );

bootstrap Win32::OLE;

# compatibility

*Win32::OLELastError = \&Win32::OLE::LastError;
*Win32::OLECreateObject = \&Win32::OLE::CreateObject;
*Win32::OLEDestroyObject = \&Win32::OLE::DestroyObject;
*Win32::OLEDispatch = \&Win32::OLE::Dispatch;
*Win32::OLEGetProperty = \&Win32::OLE::GetProperty;
*Win32::OLESetProperty = \&Win32::OLE::SetProperty;

# helper routines. see ole.xs for all the gory stuff.

sub AUTOLOAD {
    my( $self ) = shift;
    my $fReturn = "";
    $AUTOLOAD =~ s/.*:://;
    if ( Win32::OLE::Dispatch( $self, $AUTOLOAD, $fReturn, @_ ) ) {
        return $fReturn;
    } else {
        return undef;
    }
}


# Automation data types.

sub VT_EMPTY {0;}
sub VT_NULL {1;}
sub VT_I2 {2;}
sub VT_I4 {3;}
sub VT_R4 {4;}
sub VT_R8 {5;}
sub VT_CY {6;}
sub VT_DATE {7;}
sub VT_BSTR {8;}
sub VT_DISPATCH {9;}
sub VT_ERROR {10;}
sub VT_BOOL {11;}
sub VT_VARIANT {12;}
sub VT_UNKNOWN {13;}
sub VT_I1 {16;}
sub VT_UI1 {17;}
sub VT_UI2 {18;}
sub VT_UI4 {19;}
sub VT_I8 {20;}
sub VT_UI8 {21;}
sub VT_INT {22;}
sub VT_UINT {23;}
sub VT_VOID {24;}
sub VT_HRESULT {25;}
sub VT_PTR {26;}
sub VT_SAFEARRAY {27;}
sub VT_CARRAY {28;}
sub VT_USERDEFINED {29;}
sub VT_LPSTR {30;}
sub VT_LPWSTR {31;}
sub VT_FILETIME {64;}
sub VT_BLOB {65;}
sub VT_STREAM {66;}
sub VT_STORAGE {67;}
sub VT_STREAMED_OBJECT {68;}
sub VT_STORED_OBJECT {69;}
sub VT_BLOB_OBJECT {70;}
sub VT_CF {71;}
sub VT_CLSID {72;}


# Typelib

sub TKIND_ENUM {0;}
sub TKIND_RECORD {1;}
sub TKIND_MODULE {2;}
sub TKIND_INTERFACE {3;}
sub TKIND_DISPATCH {4;}
sub TKIND_COCLASS {5;}
sub TKIND_ALIAS {6;}
sub TKIND_UNION {7;}
sub TKIND_MAX {8;}

sub new {
    my( $object );
    my( $c ) = shift;
    my( $type ) = shift;
    if ( CreateObject( $type, $object ) ) {
        return $object;
    } else {
        return undef;
    }
}

sub Variant {
    return Win32::OLE::Variant->new(@_);
}

#sub DESTROY {
#    my( $self ) = shift;
#    #warn "Destroy called on $self\n";
#}

# egregium
*OLECreateObject = \&new;

package Win32::OLE::Variant;

sub new {
    my $self = {};
    my $pack = shift;
    $self->{'Type'} = shift;
    $self->{'Value'} = shift;
    return bless $self, $pack;
}

1;

__END__

=head1 NAME

Win32::OLE - OLE Automation extensions and Variants

=head1 SYNOPSIS

	$ex = new Win32::OLE 'Excel.Application' or die "oops\n";
	$ex->Amethod("arg")->Bmethod->{'Property'} = "foo";

=head1 DESCRIPTION

This module provides an interface to OLE Automation from Perl.
OLE Automation brings VisualBasic like scripting capabilities and
offers powerful extensibility and the ability to control many Win32
applications from Perl scripts.

OCX's are currently not supported.

=head2 Functions/Methods

=over 8

=item new Win32::OLE $oleclass

OLE Automation objects are created using the new() method, the
second argument to which must be the OLE class of the application
to create.  Return value is undef if the attempt to create an
OLE connection failed for some reason.

The object returned by the new() method can be used to invoke
methods or retrieve properties in the same fashion as described
in the documentation for the particular OLE class (eg. Microsoft
Excel documentation describes the object hierarchy along with the
properties and methods exposed for OLE access).

Properties can be retrieved or set using hash syntax, while methods
can be invoked with the usual perl method call syntax.

If a method or property returns an embedded OLE object, method
and property access can be chained as shown in the examples below.

=item Variant(TYPENAME, DATA)

This function returns a Win32::OLE::Variant object of the specified
type that contains the given data.  The Win32::OLE::Variant object
can be used to specify data types other than IV, NV or PV (which are
supported transparently).  See L<Variants> below for details.

=back

=head2 Constants

These constants are exported by default:

	VT_UI1
	VT_I2
	VT_I4
	VT_R4
	VT_R8
	VT_DATE
	VT_BSTR
	VT_CY
	VT_BOOL

Other OLE constants are also defined in the Win32::OLE package,
but they are unsupported at this time, so they are exported
only on request:

	VT_EMPTY
	VT_NULL
	VT_DISPATCH
	VT_ERROR
	VT_VARIANT
	VT_UNKNOWN
	VT_UI2
	VT_UI4
	VT_I8
	VT_UI8
	VT_INT
	VT_UINT
	VT_VOID
	VT_HRESULT
	VT_PTR
	VT_SAFEARRAY
	VT_CARRAY
	VT_USERDEFINED
	VT_LPSTR
	VT_LPWSTR
	VT_FILETIME
	VT_BLOB
	VT_STREAM
	VT_STORAGE
	VT_STREAMED_OBJECT
	VT_STORED_OBJECT
	VT_BLOB_OBJECT
	VT_CF
	VT_CLSID

=head2 Variants

A Variant is a data type that is used to pass data between OLE
connections.

The default behavior is to convert each perl scalar variable into
an OLE Variant according to the internal perl representation.
The following type correspondence holds:

        C type          Perl type       OLE type
        ------          ---------       --------
          int              IV            VT_I4
        double             NV            VT_R8
        char *             PV            VT_BSTR

Note that VT_BSTR is a wide character or Unicode string.  This presents a
problem if you want to pass in binary data as a parameter as 0x00 is
inserted between all the bytes in your data. The C<Variant()> method
provides a solution to this.  With Variants the script
writer can specify the OLE variant type that the parameter should be
converted to.  Currently supported types are:

        VT_UI1     unsigned char
        VT_I2      signed int (2 bytes)
        VT_I4      signed int (4 bytes)
        VT_R4      float      (4 bytes)
        VT_R8      float      (8 bytes)
        VT_DATE    OLE Date
        VT_BSTR    OLE String
        VT_CY      OLE Currency
        VT_BOOL    OLE Boolean

When VT_DATE and VT_CY objects are created, the input
parameter is treated as a Perl string type, which is then converted
to VT_BSTR, and finally to VT_DATE of VT_CY using the VariantChangeType()
OLE API function.  See L<EXAMPLES> for how these types can be used.

=head1 EXAMPLES

Here is a simple Microsoft Excel application.

	use Win32::OLE;
	$ex = new Win32::OLE 'Excel.Application' or die "oops\n";
	
	# open an existing workbook
	$ex->Workbooks->Open( 'test.xls' );
	
	# write to a particular cell
	$ex->Workbooks(1)->Worksheets('Sheet1')->Cells(1,1)->{Value} = "foo";
	
	# save and exit
	$ex->Save;
	$ex->Quit;

Here is an example of using Variant data types.

	use Win32::OLE;
	$ex = new Win32::OLE 'Excel.Application' or die "oops\n";
	$ex->{Visible} = 1;
	$ex->Workbooks->Add;
	$ovR8 = Variant(VT_R8, "3 is a good number");
	$ex->Range("A1")->{Value} = $ovR8;
	$ex->Range("A2")->{Value} = Variant(VT_DATE, 'Jan 1,1970');

The above will put value "3" in cell A1 rather than the string
"3 is a good number".  Cell A2 will contain the date.

Similarly, to invoke a method with some binary data, you can
do the following:

	$obj->Method( Variant(VT_UI1, "foo\000b\001a\002r") );

Here is a wrapper class that basically delegates everything but
new() and DESTROY().  Such a wrapper is needed for properly
shutting down connections if your application is liable to
die without proper cleanup.

	package Excel;
	use Win32::OLE;
	
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
		$s->{Ex}->Quit;
		return undef;
	    }
	}
	
	sub AUTOLOAD {
	    my $s = shift;
	    $AUTOLOAD =~ s/^.*:://;
	    $s->{Ex}->$AUTOLOAD(@_);
	}
	
	1;

The above module can be used just like Win32::OLE, except that
it takes care of closing connections in case of abnormal exits.

=head1 NOTES

There are some incompatibilities with the version distributed by Activeware
(as of build 306).

=over 4

=item 1

The package name has changed from "OLE" to "Win32::OLE".

=item 2

All functions of the form "Win32::OLEFoo" are now "Win32::OLE::Foo",
though the old names are temporarily accomodated.

=item 3

Package "OLE::Variant" is now "Win32::OLE::Variant".

=item 4

The Variant function is new, and is exported by default.  So are
all the VT_XXX type constants.

=back

You are responsible for properly closing any open OLE servers
down.  For example, if you open a OLE connection to Excel and
subsequently just die(), Excel will not shutdown and you will have
a process leak on your hands.  You will need to wrap the OLE
connection in your own object and provide a DESTROY method that
does proper cleanup to ensure smooth shutdown.  Alternatively,
you can use a __DIE__ hook or an END{} block to do such cleanup.
See L<EXAMPLES> above for an example of using a wrapper object.

=head1 AUTHORS

Originally put together by the kind people at Hip and Activeware.

Gurusamy Sarathy <gsar@umich.edu> has subsequently fixed several
major bugs, memory leaks, and reliability problems, along with some
redesign of the code.

=head1 COPYRIGHT

    (c) 1995 Microsoft Corporation. All rights reserved. 
	Developed by ActiveWare Internet Corp., http://www.ActiveWare.com

    Other modifications (c) 1997 by Gurusamy Sarathy <gsar@umich.edu>

    You may distribute under the terms of either the GNU General Public
    License or the Artistic License, as specified in the README file.


=cut


