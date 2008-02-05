package Win32::OLE;

# The documentation is at the __END__

use strict;
use vars qw($VERSION @ISA @EXPORT @EXPORT_OK $AUTOLOAD);

$VERSION = '0.05';

# Do not "use Carp;", it pollutes the OLE namespace!
# It must be required though, because the XS code uses
# Carp::croak for error reporting!
require Carp;

# We import the variables from Win32::OLE::Variant before we export them again.
# In the next version the user must "use Win32::OLE::Variant" directly.
use Win32::OLE::Variant qw(!new !Variant);

use Exporter;
use DynaLoader;
@ISA = qw(Exporter DynaLoader);

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
		With
	       );

bootstrap Win32::OLE;

# The following class methods are pure XS code. They will delegate
# to Dispatch when called as object methods.
#
# - new(oleclass,destroy)
# - LastError()
# - GetActiveObject(oleclass)
# - GetObject(pathname)
# - QueryObjectType(object)
#
# The following method is pure XS (and not available as OLE method)
# - DESTROY()
#

# <compatibility> (deprecated, will be gone in next version)
sub GetProperty {
    Carp::carp("Use of Win32::OLE::GetProperty is deprecated") if $^W;
    # GetProperty($object,$varName,$varReturn)
    local $^W = 1;
    eval { $_[2] = $_[0]->{$_[1]}; };
    return !$@;
}

sub SetProperty {
    Carp::carp("Use of Win32::OLE::SetProperty is deprecated") if $^W;
    # SetProperty($object,$varName,$varValue)
    local $^W = 1;
    eval { $_[0]->{$_[1]} = $_[2]; };
    return !$@;
}

*OLECreateObject = \&new;
*Win32::OLELastError = \&Win32::OLE::LastError;
*Win32::OLECreateObject = \&Win32::OLE::CreateObject;
#*Win32::OLEDestroyObject = \&Win32::OLE::DestroyObject;
*Win32::OLEDispatch = \&Win32::OLE::Dispatch;
*Win32::OLEGetProperty = \&Win32::OLE::GetProperty;
*Win32::OLESetProperty = \&Win32::OLE::SetProperty;

# </compatibility>

sub CreateObject {
    if (ref($_[0]) && UNIVERSAL::isa($_[0],'Win32::OLE')) {
	$AUTOLOAD = 'CreateObject';
	goto &AUTOLOAD;
    }
    # $Success = Win32::OLE->CreateObject($Class,$Object);
    $_[1] = Win32::OLE->new($_[0]);
    return defined $_[1];
}

sub AUTOLOAD {
    my $self = shift;
    my $retval;
    $AUTOLOAD =~ s/.*:://;
    Carp::croak("Cannot autoload class method \"$AUTOLOAD\"") 
      unless ref($self) && UNIVERSAL::isa($self,'Win32::OLE');
    $self->Dispatch($AUTOLOAD, $retval, @_);
    return $retval;
}

sub Variant {
    if (ref($_[0]) && UNIVERSAL::isa($_[0],'Win32::OLE')) {
	$AUTOLOAD = 'Variant';
	goto &AUTOLOAD;
    }
    return Win32::OLE::Variant->new(@_);
}

sub With {
    my $object = shift;
    while (@_) {
	my $property = shift;
	$object->{$property} = shift;
    }
}

1;

__END__

=head1 NAME

Win32::OLE - OLE Automation extensions

=head1 SYNOPSIS

    $ex = Win32::OLE->new('Excel.Application') or die "oops\n";
    $ex->Amethod("arg")->Bmethod->{'Property'} = "foo";
    $ex->Cmethod(undef,undef,$Arg3);
    $ex->Dmethod($RequiredArg1, {NamedArg1 => $Value1, NamedArg2 => $Value2});

    $wd = Win32::OLE->GetObject("D:\\Data\\Message.doc");
    $xl = Win32::OLE->GetActiveObject("Excel.Application");

=head1 DESCRIPTION

This module provides an interface to OLE Automation from Perl.
OLE Automation brings VisualBasic like scripting capabilities and
offers powerful extensibility and the ability to control many Win32
applications from Perl scripts.

OLE events and OCX's are currently not supported.

=head2 Functions/Methods

=over 8

=item Win32::OLE->new(CLASS [, DESTRUCTOR])

OLE Automation objects are created using the new() method, the
second argument to which must be the OLE class of the application
to create.  Return value is undef if the attempt to create an
OLE connection failed for some reason. The optional third argument
specifies a DESTROY-like method. This can be either a CODE reference
or a string containing an OLE method name. It can be used to cleanly
terminate OLE objects in case the Perl program dies in the middle of
OLE activity.

The object returned by the new() method can be used to invoke
methods or retrieve properties in the same fashion as described
in the documentation for the particular OLE class (eg. Microsoft
Excel documentation describes the object hierarchy along with the
properties and methods exposed for OLE access).

Optional parameters on method calls can be omitted by using C<undef>
as a placeholder. A better way is to use named arguments, as the
order of optional parameters may change in later versions of the OLE
server application. Named parameters can be specified in a reference
to a hash as the last parameter to a method call.

Properties can be retrieved or set using hash syntax, while methods
can be invoked with the usual perl method call syntax. The C<keys>
and C<each> functions can be used to enumerate an object's properties.
Beware that a property is not always writable or even readable (sometimes
raising exceptions when read while being undefined).

If a method or property returns an embedded OLE object, method
and property access can be chained as shown in the examples below.

=item Win32::OLE->GetActiveObject(CLASS)

The GetActiveObject class method returns an OLE reference to a
running instance of the specified OLE automation server. It returns
C<undef> if the server is not currently active. It will croak if
the class is not even registered.

=item Win32::OLE->GetObject(MONIKER)

The GetObject class method returns an OLE reference to the specified
object. The object is specified by a pathname optionally followed by
additional item subcomponent separated by exclamation marks '!'.

=item Win32::OLE->QueryObjectType(OBJECT)

The QueryObjectType class method returns a list of the type library
name and the objects class name. In a scalar context it returns the
class name only.

=item With(OBJECT, PROPERTYNAME => VALUE, ...)

This utility function provides a concise way to set multiple properties
on an object.  It iterates over its arguments doing
C<$OBJECT->{PROPERTYNAME} = $VALUE> on each trailing pair.  This
function is not exported by default.

=back

=head1 EXAMPLES

Here is a simple Microsoft Excel application.

	use Win32::OLE;

	# use existing instance if Excel is already running
	eval {$ex = Win32::OLE->GetActiveObject('Excel.Application')};
	if ($@) {
	    $ex = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
		    or die "oops\n";
	}
	
	# open an existing workbook
	$book = $ex->Workbooks->Open( 'test.xls' );
	
	# write to a particular cell
	$book->Worksheets(1)->Cells(1,1)->{Value} = "foo";
	
	# save and exit
	$book->Save;
	undef $book;
	undef $ex;

Please note the destructor specified on the Win32::OLE->new method. It ensures
that Excel will shutdown properly even if the Perl program dies. Otherwise
there could be a process leak if your application dies after having opened
an OLE instance of Excel. It is the responsibility of the module user to
make sure that all OLE objects are cleaned up properly!

Here is an example of using Variant data types.

	use Win32::OLE;
	$ex = Win32::OLE->new('Excel.Application', \&OleQuit) or die "oops\n";
	$ex->{Visible} = 1;
	$ex->Workbooks->Add;
	$ovR8 = Variant(VT_R8, "3 is a good number");
	$ex->Range("A1")->{Value} = $ovR8;
	$ex->Range("A2")->{Value} = Variant(VT_DATE, 'Jan 1,1970');

	sub OleQuit { 
	    my $self = shift; 
	    $self->Quit; 
	}

The above will put value "3" in cell A1 rather than the string
"3 is a good number".  Cell A2 will contain the date.

Similarly, to invoke a method with some binary data, you can
do the following:

	$obj->Method( Variant(VT_UI1, "foo\000b\001a\002r") );

Here is a wrapper class that basically delegates everything but
new() and DESTROY().  The wrapper class shown here is another way to
properly shut down connections if your application is liable to die
without proper cleanup.  Your own wrappers will probably do something
more specific to the particular OLE object you may be dealing with,
like overriding the methods that you may wish to enhance with your
own.

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
Note that the effect of this specific example can be easier accomplished
using the optional destructor argument of Win32::OLE::new:

	my $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;});

Note that the delegation shown in the earlier example is not the same as
true subclassing with respect to further inheritance of method calls in your
specialized object.  See L<perlobj>, L<perltoot> and L<perlbot> for details.
True subclassing (available by setting C<@ISA>) is also feasible,
as the following example demonstrates:

	#
	# Add error reporting to Win32::OLE
	#
	
	package Win32::OLE::Strict;
	use Carp;
	use Win32::OLE;
	
	use strict qw(vars);
	use vars qw($AUTOLOAD @ISA);
	@ISA = qw(Win32::OLE);
	
	sub AUTOLOAD {
	    my $obj = shift;
	    $AUTOLOAD =~ s/^.*:://;
	    my $meth = $AUTOLOAD;
	    $AUTOLOAD = "SUPER::" . $AUTOLOAD;
	    my $retval = $obj->$AUTOLOAD(@_);
	    unless (defined($retval) || $AUTOLOAD eq 'DESTROY') {
		my $err = Win32::OLE::LastError();
		croak(sprintf("$meth returned OLE error 0x%08x",$err))
		  if $err;
	    }
	    return $retval;
	}
	
	1;

Here's how the above class will be used:

	use Win32::OLE::Strict;
	my $Excel = Win32::OLE::Strict->new('Excel.Application', 'Quit');
	my $Books = $Excel->Workbooks;
	$Books->UnknownMethod(42);

In the sample above the call to C<UnknownMethod> will be caught with

	UnknownMethod returned OLE error 0x80020009 at test.pl line 5

because the Workbooks object inherits the class C<Win32::OLE::Strict> from the
C<$Excel> object.

=head1 NOTES

=head2 Hints for Microsoft Office automation

=over 8

=item Documentation

The object model for the Office applications is defined in the Visual Basic
reference guides for the various applications. These are typically not
installed by default during the standard installation. They can be added
later by rerunning the setup program with the custom install option.

=item Class, Method and Property names

The names have been changed between different versions of Office. For
example C<Application> was a method in Office 95 and is a property in
Office97. Therefore it will not show up in the list of property names
C<keys %$object> when querying an Office 95 object.

The class names are not always identical to the method/property names
producing the object. E.g. the C<Workbook> method returns an object of
type C<Workbook> in Office 95 and C<_Workbook> in Office 97.

=item Moniker (GetObject support)

Office applications seem to implement file monikers only. For example
it seems to be impossible to retrieve a specific worksheet object through
C<GetObject("File.XLS!Sheet")>. Furthermore, in Excel 95 the moniker starts
a Worksheet object and in Excel 97 it returns a Workbook object. You can use
either the Win32::OLE::QueryObjectType class method or the $object->{Version}
property to write portable code.

=item Enumeration of collection objects

Enumerations seem to be incompletely implemented. Office 95 application don't
seem to support neither the Reset() nor the Clone() methods. The Clone()
method is still unimplemented in Office 97. A single walk through the
collection similar to Visual Basics C<for each> construct does work however.

=item Localization

Starting with Office 97 Microsoft has changed the localized class, method and
property names back into English. Note that string, date and currency
arguments are still subject to locale specific interpretation. Perl uses the
system default locale for all OLE transaction whereas Visual Basic uses a 
type library specific locale. A Visual Basic script would use "R1C1" in string
arguments to specify relative references. A Perl script running on a German
language Windows would have to use "Z1S1". The next version of Win32::OLE will
allow writing more portable scripts in this regard.

=item SaveAs method in Word 97 doesn't work

This is an known bug in Word 97. Search the MS knowledge base for Word /
Foxpro incompatibility. The problems applies to the Perl OLE interface as
well. A workaround is to use the WordBasic compatibility object. It doesn't
support all the options of the native method though.

    $Word->WordBasic->FileSaveAs($file);

=back

=head2 Incompatibilities

There are some incompatibilities with the version distributed by Activeware
(as of build 306).

=over 8

=item 1

The package name has changed from "OLE" to "Win32::OLE".

=item 2

All functions of the form "Win32::OLEFoo" are now "Win32::OLE::Foo",
though the old names are temporarily accomodated.  Win32::OLECreateObject()
was changed to Win32::OLE::CreateObject(), and is now called
Win32::OLE::new() bowing to established convention for naming constructors.
The old names should be considered deprecated, and will be removed in the
next version.

=item 3

Package "OLE::Variant" is now "Win32::OLE::Variant".

=item 4

The Variant function is new, and is exported by default.  So are
all the VT_XXX type constants.

=item 5

The support for collection objects has been moved into the package
Win32::OLE::Enum. The C<keys %$object> method is now used to enumerate
the properties of the object.

=back

=head2 Bugs and Limitations

=over 8

=item 1

Currently there is no way to invoke any of the C<Dispatch>, C<DESTROY>, 
C<GetProperty>, C<SetProperty> or C<With> object methods (except by
calling Dispatch directly). This will be fixed in the next release
by providing a documented replacement for the C<Dispatch> method. The
interface has not been determined yet.

=item 2

In the current release all the C<VT_*> and C<TKIND_*> names are not
available as OLE method names.

=item 3

All function names defined by the Exporter module are currently unavailable
as OLE method names. They are C<export>, C<export_to_level>, C<import>,
C<_push_tags>, C<export_tags>, C<export_ok_tags>, C<export_fail> and
C<require_version>.

The same is true for all names defined by the Dynaloader: C<dl_load_flags>,
C<croak>, C<bootstrap>, C<dl_findfile>, C<dl_expandspec>, 
C<dl_find_symbol_anywhere>, C<dl_load_file>, C<dl_find_symbol>,
C<dl_undef_symbols>, C<dl_install_xsub> and C<dl_error>.

=item 4

The implementation is rather sensitive to error conditions, and will
croak() on many different kinds of errors encountered at run time.  This
could be construed as improper behavior for a generic module such as this.
A well defined error API to report exceptional conditions will be offered
in future, to allow the user to control which conditions are fatal.

=back

=head2 Deprecated features

=over 8

=item 1

All C<Win32::OLE*> (but not C<Win32::OLE::*>) methods will be removed in the
next release. The C<GetProperty> and C<SetProperty> methods will be removed
too. They were always undocumented and at least C<SetProperty> also didn't
really seem to work anyways.

=item 2

The Variant functions and constants have been moved into the separate module
Win32::OLE::Variant. Starting with the next release the variant functions
and constants will be only available through that package anymore.

=back

=head1 AUTHORS

Originally put together by the kind people at Hip and Activeware.

Gurusamy Sarathy <gsar@umich.edu> subsequently fixed several major
bugs, memory leaks, and reliability problems, along with some
redesign of the code.

Jan Dubois <jan.dubois@ibm.net> pitched in with yet more massive redesign,
added support for named parameters, and other significant enhancements.

=head1 COPYRIGHT

    (c) 1995 Microsoft Corporation. All rights reserved. 
	Developed by ActiveWare Internet Corp., http://www.ActiveWare.com

    Other modifications (c) 1997 by Gurusamy Sarathy <gsar@umich.edu>
    and Jan Dubois <jan.dubois@ibm.net>

    You may distribute under the terms of either the GNU General Public
    License or the Artistic License, as specified in the README file.

=head1 VERSION

Version 0.05	14 December 1997

=cut


