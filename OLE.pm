# The documentation is at the __END__

package Win32::OLE;

use strict;
use vars qw($VERSION @ISA @EXPORT @EXPORT_OK @EXPORT_FAIL $AUTOLOAD
	    $CP $LCID $Warn $LastError);

$VERSION = '0.0608';

use Carp;
use Exporter;
use DynaLoader;
@ISA = qw(Exporter DynaLoader);

@EXPORT = qw();
@EXPORT_OK = qw(CP_ACP CP_OEMCP in valof with OVERLOAD);
@EXPORT_FAIL = qw(OVERLOAD);

sub export_fail {
    shift;
    if ($_[0] eq 'OVERLOAD') {
	shift;
	eval <<'OVERLOAD';
	    use overload '""'     => \&valof,
	                 '0+'     => \&valof,
	                 fallback => 1;
OVERLOAD
    }
    return @_;
}

bootstrap Win32::OLE;

$Warn = $^W;

sub CP_ACP {0;}    # ANSI codepage
sub CP_OEMCP {1;}  # OEM codepage

# The following class methods are pure XS code. They will delegate
# to Dispatch when called as object methods.
#
# - new(oleclass,destroy)
# - GetActiveObject(oleclass)
# - GetObject(pathname)
# - QueryObjectType(object)
#
# The following method is pure XS (and not available as OLE method)
# - DESTROY()
#


# CreateObject is defined here because it is documented in the
# "Learning Perl on Win32 Systems" book. Please use Win32::OLE->new().
sub CreateObject {
    if (ref($_[0]) && UNIVERSAL::isa($_[0],'Win32::OLE')) {
	$AUTOLOAD = 'CreateObject';
	goto &AUTOLOAD;
    }

    $_[1] = Win32::OLE->new($_[0]);
    return defined $_[1];
}

sub LastError {
    unless (defined $_[0]) {
	carp("LastError must be called as class method!");
	return;
    }

    if (ref($_[0]) && UNIVERSAL::isa($_[0],'Win32::OLE')) {
	$AUTOLOAD = 'LastError';
	goto &AUTOLOAD;
    }

    no strict 'refs';
    my $LastError = "$_[0]::LastError";
    $$LastError = $_[1] if defined $_[1];
    return $$LastError;
}

sub Invoke {
    my ($self, $method, @args) = @_;
    my $retval;
    $self->Dispatch($method, $retval, @args);
    return $retval;
}

sub AUTOLOAD {
    my $self = shift;
    my $retval;
    $AUTOLOAD =~ s/.*:://o;
    croak("Cannot autoload class method \"$AUTOLOAD\"") 
      unless ref($self) && UNIVERSAL::isa($self, 'Win32::OLE');
    my $success = $self->Dispatch($AUTOLOAD, $retval, @_);
    unless (defined $success || ($^H & 0x200)) {
	# Retry default method if C<no strict 'subs';>
	$self->Dispatch(undef, $retval, $AUTOLOAD, @_);
    }
    return $retval;
}

sub in {
    my @res;
    require Win32::OLE::Enum;
    while (@_) {
	my $this = shift;
	if (UNIVERSAL::isa($this, 'Win32::OLE')) {
	    push @res, Win32::OLE::Enum->All($this);
	}
	elsif (ref($this) eq 'ARRAY') {
	    push @res, @$this;
	}
	else {
	    push @res, $this;
	}
    }
    return @res;
}

sub valof {
    require Win32::OLE::Variant;
    my $arg = shift;
    if (UNIVERSAL::isa($arg, 'Win32::OLE')) {
	my ($class) = overload::StrVal($arg) =~ /^([^=]+)=/;
	no strict 'refs';
	local $Win32::OLE::Variant::CP = $ {$class."::CP"};
	local $Win32::OLE::Variant::LCID = $ {$class."::LCID"};
	use strict 'refs';
	# VT_EMPTY variant for return code
	my $variant = Win32::OLE::Variant->new(0,0);
	$arg->Dispatch(undef, $variant);
	return $variant->Value;
    }
    $arg = $arg->Value if UNIVERSAL::can($arg, 'Value');
    return $arg;
}

sub with {
    my $object = shift;
    while (@_) {
	my $property = shift;
	$object->{$property} = shift;
    }
}

########################################################################

package Win32::OLE::Tie;

# Only retry default method under C<no strict 'subs';>

sub FETCH {
    my ($self,$key) = @_;
    $self->Fetch($key, ~($^H & 0x200));
}

sub STORE {
    my ($self,$key,$value) = @_;
    $self->Store($key, $value, ~($^H & 0x200));
}

1;

########################################################################

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

The Win32::OLE module uses the IDispatch interface exclusively. It is
not possible to access a custom OLE interface. OLE events and OCX's are
currently not supported.

=head2 Methods

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

=item OBJECT->Invoke(METHOD,ARGS)

The C<Invoke> object method is an alternate way to invoke OLE
methods. It is normally equivalent to C<$OBJECT->METHOD(@ARGS)>. This
function must be used if the METHOD name contains characters not valid
in a Perl variable name (like foreign language characters). It can
also be used to invoke the default method of an object even if the
default method has not been given a name in the type library. In this
case use <undef> or C<''> as the method name.

=item Win32::OLE->LastError()

The C<LastError> class method returns the last recorded OLE
error. This is dual value like the C<$!> variable: in a numeric
context it returns the error number and in a string context it returns
the error message.

The last OLE error is not automatically reset by a successful OLE
call. The numeric value can be explicitly set by a call (which will
discard the string value):

	Win32::OLE->LastError(0);

=item Win32::OLE->QueryObjectType(OBJECT)

The QueryObjectType class method returns a list of the type library
name and the objects class name. In a scalar context it returns the
class name only. It returns C<undef> when the type information is not
available.

=back

Whenever Perl does not find a method name in the Win32::OLE package it
is automatically used as the name of an OLE method and this method call
is dispatched to the OLE server.

There is one special hack built into the module: If a method or property 
name could not be resolved with the OLE object, then the default method
of the object is called with the method name as its first parameter. So

	my $Sheet = $Worksheets->Table1;
or
	my $Sheet = $Worksheets->{Table1};

is resolved as

	my $Sheet = $Worksheet->Item('Table1');

provided that the C<$Worksheets> object doesnot have a C<Table1> method
or property. This hack has been introduced to call the default method
of collections which did not name the method in their type library. The
recommended way to call the "unnamed" default method is:

	my $Sheet = $Worksheets->Invoke('', 'Table1');

This special hack is disabled under C<use strict 'subs';>.

=head2 Functions

The following functions are not exported by default.

=over 8

=item in(COLLECTION)

If COLLECTION is an OLE collection object then C<in $COLLECTION>
returns a list of all members of the collection. This is a shortcut
for C<Win32::OLE::Enum->All($COLLECTION)>. It is most commonly used in
a C<foreach> loop:

	foreach my $value (in $collection) {
	    # do something with $value here
	}

=item valof(OBJECT)

Normal assignment of Perl OLE objects creates just another reference
to the OLE object. The C<valof> function explictly dereferences the
object (through the default method) and returns the value of the object.

	my $RefOf = $Object;
	my $ValOf = valof $Object;
        $Object->{Value} = $NewValue;

Now C<$ValOf> still contains the old value wheras C<$RefOf> would
resolve to the C<$NewValue> because it is still a reference to
C<$Object>.

The C<valof> function can also be used to convert Win32::OLE::Variant
objects to Perl values.

=item with(OBJECT, PROPERTYNAME => VALUE, ...)

This function provides a concise way to set the values of multiple
properties of an object.  It iterates over its arguments doing
C<$OBJECT->{PROPERTYNAME} = $VALUE> on each trailing pair.

=back

=head2 Overloading

The Win32::OLE objects can be overloaded to automatically convert to
their values whenever they are used in a bool, numeric or string
context. This is not enabled by default. You have to request it
through the C<OVERLOAD> pseudotarget:

	use Win32::OLE qw(in valof with OVERLOAD);

Please note that this is a global setting. If any module enables
Win32::OLE overloading then it's active everywhere.

=head2 Class Variables

=over 8

=item $Win32::OLE::CP

This variable is used to determine the codepage used by all
translations between Perl strings and Unicode strings used by the OLE
interface. The default value is CP_ACP, which is the default ANSI
codepage. It can also be set to CP_OEMCP which is the default OEM
codepage. Both constants are not exported by default.

=item $Win32::OLE::LCID

This variable controls the locale idnetifier used for all OLE calls.
It is set to LOCALE_NEUTRAL by default. Please check the
L<Win32::OLE::NLS> module for other locale related information.

=item $Win32::OLE::Warn

This variable determines the behavior of the Win32::OLE module when
an error happens. Valid values are:

	0	Ignore error, return undef
	1	Carp::carp if $^W is set (-w option)
	2	always Carp::carp
	3	Carp::croak

The error number and message (without Carp line/module info) are
available through the C<Win32::OLE->LastError> class method.

=back

=head1 EXAMPLES

Here is a simple Microsoft Excel application.

	use Win32::OLE;

	# use existing instance if Excel is already running
	eval {$ex = Win32::OLE->GetActiveObject('Excel.Application')};
	die "Excel not installed" if $@;
	unless (defined $ex) {
	    $ex = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
		    or die "Oops, cannot start Excel";
	}
	
	# open an existing workbook
	$book = $ex->Workbooks->Open( 'test.xls' );
	
	# write to a particular cell
	$sheet = $book->Worksheets(1);
	$sheet->Cells(1,1)->{Value} = "foo";

        # write a 2 rows by 3 columns range
        $sheet->Range("A8:C9")->{Value} = [[ undef, 'Xyzzy', 'Plugh' ],
                                           [ 42,    'Perl',  3.1415  ]];

        # print "XyzzyPerl"
        $array = $sheet->Range("A8:B9")->{Value};
        print $array[0][1] . $array[1][1];

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

This package inherits the constructor C<new()> from the Win32::OLE
package. It is important to note that you cannot later rebless a
Win32::OLE object as some information about the package is cached by
the object. Always invoke the C<new()> constructor through the right
package!

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
language Windows would have to use "Z1S1". Set the C<$Win32::OLE::LCID> class
variable to an English locale to write portable scripts. This variable should
not be changed after creating the OLE objects; some methods seem to randomly
fail if the locale is changed on the fly.

=item SaveAs method in Word 97 doesn't work

This is an known bug in Word 97. Search the MS knowledge base for Word /
Foxpro incompatibility. That problem applies to the Perl OLE interface as
well. A workaround is to use the WordBasic compatibility object. It doesn't
support all the options of the native method though.

    $Word->WordBasic->FileSaveAs($file);

The problem seems to be fixed by applying the Office 97 Service Release 1.

=item Randomly failing method calls

It seems like modifying objects that are not selected/activated is sometimes
fragile. Most of these problems go away if the chart/sheet/document is
selected or activated before being manipulated (just like an interactive
user would automatically do it).

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

    Other modifications Copyright (c) 1997, 1998 by Gurusamy Sarathy
    <gsar@umich.edu> and Jan Dubois <jan.dubois@ibm.net>

    You may distribute under the terms of either the GNU General Public
    License or the Artistic License, as specified in the README file.

=head1 VERSION

Version 0.06	6 February 1998

=cut


