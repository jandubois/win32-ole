package Win32::OLE::Const;

# The documentation is at the __END__

use strict;
use Carp;
use Win32::OLE;
use Win32::Registry;

sub import {
    my ($self,$name,$major,$minor,$language) = @_;
    return unless defined($name) && $name !~ /^\s*$/;
    my $callpkg = caller(0);

    my $const = $self->Load($name,$major,$minor,$language);
    while (defined(my $key = each %$const)) {
	# export only valid variable names
	next unless $key =~ /^[a-zA-Z_][a-zA-Z0-9_]*$/;
	no strict 'refs';
	*{"${callpkg}::${key}"} = \$const->{$key};
    }
}

sub Load {
    my ($pack,$name,$major,$minor,$language) = @_;
    undef $minor unless defined $major;

    return _Load($name,undef,undef,undef,undef)
      if UNIVERSAL::isa($name,'Win32::OLE');

    my ($hTypelib,$hClsid,$hVersion,$hLangid);
    my @found;

    $main::HKEY_CLASSES_ROOT->Create('TypeLib',$hTypelib) 
      or croak "Cannot access HKEY_CLASSES_ROOT\\Typelib";
    my $Clsids = [];
    $hTypelib->GetKeys($Clsids);

    foreach my $clsid (@$Clsids) {
	$hTypelib->Create($clsid,$hClsid);
	my $Versions = [];
	$hClsid->GetKeys($Versions);
	foreach my $version (@$Versions) {
	    my $value;
	    next unless $hClsid->QueryValue($version,$value);
	    next unless $value =~ /^$name/;

	    my ($maj,$min) = ($version =~ /^(\d+)\.(\d+)$/);
	    next unless defined $min;
	    next if defined($major) && $maj != $major;
	    next if defined($minor) && $min < $minor;

	    $hClsid->Create($version,$hVersion);
	    my $Langids = [];
	    $hVersion->GetKeys($Langids);
	    foreach my $langid (@$Langids) {
		next unless $langid =~ /^\d+$/;
		next if defined($language) && $language != $langid;
		$hVersion->Create($langid,$hLangid);
		my $filename;
		$hLangid->QueryValue('win32',$filename);
		$hLangid->Close;
		push @found, [$clsid,$maj,$min,$langid,$filename];
	    }
	    $hVersion->Close;
	}
	$hClsid->Close;
    }
    $hTypelib->Close;

    unless (@found) {
	carp "No type library matching \"$name\" found";
	return;
    }

    @found = sort {
	# Prefer greater version number
	my $res = $b->[1] <=> $a->[1];
	$res = $b->[2] <=> $a->[2] if $res == 0;
	# Prefer default language for equal version numbers
	$res = -1 if $res == 0 && $a->[3] == 0;
	$res =  1 if $res == 0 && $b->[3] == 0;
	return $res;
    } @found;

    #printf "Loading %s\n", join(' ', @{$found[0]});
    return _Load(@{$found[0]});
}

1;

__END__

=head1 NAME

Win32::OLE::Const - Extract constant definitions from TypeLib

=head1 SYNOPSIS

    use Win32::OLE::Const ("Microsoft Excel");
    print "xlMarkerStyleDot = $xlMarkerStyleDot\n";

    my $wd = Win32::OLE::Const->Load("Microsoft Word 8\\.0 Object Library");
    foreach my $key (keys %$wd) {
        printf "$key = %s\n", $wd->{$key};
    }

=head1 DESCRIPTION

This modules makes all constants from a registered OLE type library
available to the Perl program. The constant definitions can be
imported as scalar variables providing compile time name checking.
Alternatively the constants can be returned in a hash reference
which avoids predefining various variables of unknown names and values.

=head2 Functions/Methods

=over 4

=item use Win32::OLE::Const

The C<use> statement can be used to directly import the constant names
and values into the users namespace.

    use Win32::OLE::Const (TYPELIB,MAJOR,MINOR,LANGUAGE);

The TYPELIB argument specifies a regular expression for searching
through the registry for the type library. Note that this argument is
implicitly prefixed with C<^> to speed up matches in the most common
cases. Use a typelib name like ".*Excel" to match anywhere within the
description. TYPELIB is the only required argument.

The MAJOR and MINOR arguments specify the requested version of
the type specification. If the MAJOR argument is used then only
typelibs with exactly this major version number will be matched. The
MINOR argument however specifies the minimum acceptable minor version.
MINOR is ignored if MAJOR is undefined.

If the LANGUAGE argument is used then only typelibs with exactly this
language id will be matched.

The module will select the typelib with the highest version number
satisfying the request. If no language id is specified then a the default
language (0) will be preferred over the others.

Note that only constants with valid Perl variable names will be exported,
i.e. names matching this regexp: C</^[a-zA-Z_][a-zA-Z0-9_]*$/>.

=item Win32::OLE::Const->Load

The Win32::OLE::Const->Load method returns a reference to a hash of
constant definitions.

    my $const = Win32::OLE::Const->Load(TYPELIB,MAJOR,MINOR,LANGUAGE);

The parameters are the same as for the C<use> case.

This method is generally preferrable when the typelib uses a non-english
language and the constant names contain locale specific characters not
allowed in Perl variable names.

Another advantage is that all available constants can now be enumerated.

The load method also accepts an OLE object as a parameter. In this case
the OLE object is queried about its containing type library and no registry
search is done at all. Interestingly this seems to be slower thought.

=back

=head1 EXAMPLES

The first example imports all Excel constants names into the main namespace
and prints the value of xlMarkerStyleDot (-4118).

    use Win32::OLE::Const ('Microsoft Excel 8.0 Object Library');
    print "xlMarkerStyleDot = $xlMarkerStyleDot\n";

The second example returns all Word constants in a hash ref.

    use Win32::OLE::Const;
    my $wd = Win32::OLE::Const->Load("Microsoft Word 8.0 Object Library");
    foreach my $key (keys %$wd) {
        printf "$key = %s\n", $wd->{$key};
    }
    printf "wdGreen = %s\n", $wd->{wdGreen};

The last example uses an OLE object to specify the type library:

    use Win32::OLE;
    use Win32::OLE::Const;
    my $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;});
    my $xl = Win32::OLE::Const->Load($Excel);


=head1 AUTHORS/COPYRIGHT

This module is part of the Win32::OLE distribution.

=cut
