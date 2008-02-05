# The documentation is at the __END__

package Win32::OLE::Variant;
require Win32::OLE;  # Make sure the XS bootstrap has been called

use strict;
use vars qw(@ISA @EXPORT @EXPORT_OK $CP $LCID $LastError $Warn);

use Exporter;
@ISA = qw(Exporter);

@EXPORT = qw(
	     Variant
	     VT_EMPTY VT_NULL VT_I2 VT_I4 VT_R4 VT_R8 VT_CY VT_DATE VT_BSTR
	     VT_DISPATCH VT_ERROR VT_BOOL VT_VARIANT VT_UNKNOWN VT_UI1
	     VT_BYREF
	    );

@EXPORT_OK = qw(CP_ACP CP_OEMCP);

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
sub VT_UI1 {17;}

sub VT_BYREF {0x4000;}

# Codepages
sub CP_ACP {0;}
sub CP_OEMCP {1;}

# Package variables
$Warn = $^W;
$CP   = CP_ACP;
$LCID = 2 << 10; # LOCALE_SYSTEM_DEFAULT

# following subs are pure XS code:
# - new(type,data)
# - As(type)
# - ChangeType(type)
# - Unicode

use overload '""'     => sub {$_[0]->As(VT_BSTR)},
             '0+'     => sub {$_[0]->As(VT_R8)},
             fallback => 1; 

sub LastError {
    no strict 'refs';
    my $LastError = "$_[0]::LastError";
    $$LastError = $_[1] if defined $_[1];
    return $$LastError;
}

sub Variant {
    return Win32::OLE::Variant->new(@_);
}

1;

__END__

=head1 NAME

Win32::OLE::Variant - Create and modify OLE VARIANT variables

=head1 SYNOPSIS

	use Win32::OLE::Variant;
	my $var = Variant(VT_DATE, 'Jan 1,1970');
	$OleObject->{value} = $var;
	$OleObject->Method($var);


=head1 DESCRIPTION

The IDispatch interface used by the Perl OLE module uses a universal
argument type called VARIANT. This is basically an object containing
a data type and the actual data value. The data type is specified by
the VT_xxx constants.

=head2 Methods

=over 8

=item new(TYPE, DATA)

This method returns a Win32::OLE::Variant object of the specified
type that contains the given data.  The Win32::OLE::Variant object
can be used to specify data types other than IV, NV or PV (which are
supported transparently).  See L<Variants> below for details.

=item As(TYPE)

C<As> converts the VARIANT to the new type before converting to a
Perl value. This take the current LCID setting into account. For
example a string might contain a ',' as the decimal point character.
Using C<$variant->As(VT_R8)> will correctly return the floating
point value.

The underlying variant object is NOT changed by this method.

=item ChangeType(TYPE)

This method changes the type of the contained VARIANT in place. It
returns the object itself, not the converted value.

=item LastError()

The C<LastError> method returns the last recorded OLE error in the
Win32::OLE::Variant class. This is dual value like the C<$!> variable:
in a numeric context it returns the error number and in a string
context it returns the error message.

The method corresponds to the C<Win32::OLE->LastError> method for
Win32::OLE objects.

=item Type()

The C<Type> method returns the type of the contained VARIANT.

=item Unicode()

The C<Unicode> method returns a C<Unicode::String> object. This contains
the BSTR value of the variant in network byte order. If the variant is
not currently in VT_BSTR format then a VT_BSTR copy will be produced first.

=item Value()

The C<Value> method returns the value of the VARIANT as a Perl value. The
conversion is performed in the same manner as all return values of
Win32::OLE method calls are converted.

=back

=head2 Functions

=over 8

=item Variant(TYPE, DATA)

This is just a function alias of the Win32::OLE::Variant->new()
method. This function is exported by default.

=back

=head2 Overloading

The Win32::OLE::Variant package has overloaded the conversion to
string an number formats. Therefore variant objects can be used in
arithmetic and string operations without applying the C<Value> 
method first.

=head2 Class Variables

This module supports the C<$CP>, C<$LCID> and C<$Warn> class variables.
They have the same meaning as the variables in L<Win32::OLE> of the
same name.

=head2 Constants

These constants are exported by default:

	VT_EMPTY
	VT_NULL
	VT_I2
	VT_I4
	VT_R4
	VT_R8
	VT_CY
	VT_DATE
	VT_BSTR
	VT_DISPATCH
	VT_ERROR
	VT_BOOL
	VT_VARIANT
	VT_UNKNOWN
	VT_UI1
	VT_BYREF

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
        void *           ref to AV       VT_ARRAY
           ?              undef          VT_ERROR
           ?        Win32::OLE object    VT_DISPATCH

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
OLE API function.  See L<Win32::OLE/EXAMPLES> for how these types
can be used.

=head2 Variants by reference

Some OLE servers expect parameters passed by reference so that they
can be changed in the method call. This allows methods to easily
return multiple values. There is preliminary support for this in
the Win32::OLE::Variant module:

	my $x = Variant(VT_I4|VT_BYREF, 0);
	my $y = Variant(VT_I4|VT_BYREF, 0);
	$Corel->GetSize($x, $y);
	print "Size is $x by $y\n";

After the C<GetSize> method call C<$x> and C<$y> will be set to
the respective sizes. They will still be variants. In the print
statement the overloading converts them to string representation
automatically.

This currently works for integer, number and BSTR variants. Don't
try it with VT_DISPATCH or array variants (yet).

=head1 AUTHORS/COPYRIGHT

This module is part of the Win32::OLE distribution.

=cut
