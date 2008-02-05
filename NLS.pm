# The documentation is at the __END__

package Win32::OLE::NLS;
require Win32::OLE;  # Make sure the XS bootstrap has been called

use strict;
use vars qw(@EXPORT @EXPORT_OK %EXPORT_TAGS @ISA);

use Exporter;
@ISA = qw(Exporter);

@EXPORT = qw(
	     CompareString
	     LCMapString
	     GetLocaleInfo
	     GetSystemDefaultLangID
	     GetSystemDefaultLCID
	     GetUserDefaultLangID
	     GetUserDefaultLCID

	     MAKELANGID
	     PRIMARYLANGID
	     SUBLANGID
	     MAKELCID
	     LANGIDFROMLCID
	    );

%EXPORT_TAGS = 
(
 CT	 => [qw(CT_CTYPE1 CT_CTYPE2 CT_CTYPE3)],
 C1	 => [qw(C1_UPPER C1_LOWER C1_DIGIT C1_SPACE C1_PUNCT
		C1_CNTRL C1_BLANK C1_XDIGIT C1_ALPHA)],
 C2	 => [qw(C2_LEFTTORIGHT C2_RIGHTTOLEFT C2_EUROPENUMBER
		C2_EUROPESEPARATOR C2_EUROPETERMINATOR C2_ARABICNUMBER
		C2_COMMONSEPARATOR C2_BLOCKSEPARATOR C2_SEGMENTSEPARATOR
		C2_WHITESPACE C2_OTHERNEUTRAL C2_NOTAPPLICABLE)],
 C3	 => [qw(C3_NONSPACING C3_DIACRITIC C3_VOWELMARK C3_SYMBOL C3_KATAKANA
		C3_HIRAGANA C3_HALFWIDTH C3_FULLWIDTH C3_IDEOGRAPH C3_KASHIDA
		C3_ALPHA C3_NOTAPPLICABLE)],
 NORM	 => [qw(NORM_IGNORECASE NORM_IGNORENONSPACE NORM_IGNORESYMBOLS
		NORM_IGNOREWIDTH NORM_IGNOREKANATYPE NORM_IGNOREKASHIDA)],
 LCMAP	 => [qw(LCMAP_LOWERCASE LCMAP_UPPERCASE LCMAP_SORTKEY LCMAP_HALFWIDTH
		LCMAP_FULLWIDTH LCMAP_HIRAGANA LCMAP_KATAKANA)],
 LANG	 => [qw(LANG_NEUTRAL LANG_ALBANIAN LANG_ARABIC LANG_BAHASA
		LANG_BULGARIAN LANG_CATALAN LANG_CHINESE LANG_CZECH LANG_DANISH
		LANG_DUTCH LANG_ENGLISH LANG_FINNISH LANG_FRENCH LANG_GERMAN
		LANG_GREEK LANG_HEBREW LANG_HUNGARIAN LANG_ICELANDIC
		LANG_ITALIAN LANG_JAPANESE LANG_KOREAN LANG_NORWEGIAN
		LANG_POLISH LANG_PORTUGUESE LANG_RHAETO_ROMAN LANG_ROMANIAN
		LANG_RUSSIAN LANG_SERBO_CROATIAN LANG_SLOVAK LANG_SPANISH 
		LANG_SWEDISH LANG_THAI LANG_TURKISH LANG_URDU)],
 SUBLANG => [qw(SUBLANG_NEUTRAL SUBLANG_DEFAULT SUBLANG_SYS_DEFAULT
		SUBLANG_CHINESE_SIMPLIFIED SUBLANG_CHINESE_TRADITIONAL
		SUBLANG_DUTCH SUBLANG_DUTCH_BELGIAN SUBLANG_ENGLISH_US
		SUBLANG_ENGLISH_UK SUBLANG_ENGLISH_AUS SUBLANG_ENGLISH_CAN
		SUBLANG_ENGLISH_NZ SUBLANG_ENGLISH_EIRE SUBLANG_FRENCH
		SUBLANG_FRENCH_BELGIAN SUBLANG_FRENCH_CANADIAN
		SUBLANG_FRENCH_SWISS SUBLANG_GERMAN SUBLANG_GERMAN_SWISS
		SUBLANG_GERMAN_AUSTRIAN SUBLANG_ITALIAN SUBLANG_ITALIAN_SWISS
		SUBLANG_NORWEGIAN_BOKMAL SUBLANG_NORWEGIAN_NYNORSK
		SUBLANG_PORTUGUESE SUBLANG_PORTUGUESE_BRAZILIAN
		SUBLANG_SERBO_CROATIAN_CYRILLIC SUBLANG_SERBO_CROATIAN_LATIN
		SUBLANG_SPANISH SUBLANG_SPANISH_MEXICAN
		SUBLANG_SPANISH_MODERN)],
 CTRY	 => [qw(CTRY_DEFAULT CTRY_AUSTRALIA CTRY_AUSTRIA CTRY_BELGIUM
		CTRY_BRAZIL CTRY_CANADA CTRY_DENMARK CTRY_FINLAND CTRY_FRANCE
		CTRY_GERMANY CTRY_ICELAND CTRY_IRELAND CTRY_ITALY CTRY_JAPAN
		CTRY_MEXICO CTRY_NETHERLANDS CTRY_NEW_ZEALAND CTRY_NORWAY
		CTRY_PORTUGAL CTRY_PRCHINA CTRY_SOUTH_KOREA CTRY_SPAIN
		CTRY_SWEDEN CTRY_SWITZERLAND CTRY_TAIWAN CTRY_UNITED_KINGDOM
		CTRY_UNITED_STATES)],
 LOCALE	 => [qw(LOCALE_NOUSEROVERRIDE LOCALE_ILANGUAGE LOCALE_SLANGUAGE
		LOCALE_SENGLANGUAGE LOCALE_SABBREVLANGNAME
		LOCALE_SNATIVELANGNAME LOCALE_ICOUNTRY LOCALE_SCOUNTRY
		LOCALE_SENGCOUNTRY LOCALE_SABBREVCTRYNAME LOCALE_SNATIVECTRYNAME
		LOCALE_IDEFAULTLANGUAGE LOCALE_IDEFAULTCOUNTRY
		LOCALE_IDEFAULTCODEPAGE LOCALE_IDEFAULTANSICODEPAGE LOCALE_SLIST
		LOCALE_IMEASURE LOCALE_SDECIMAL LOCALE_STHOUSAND
		LOCALE_SGROUPING LOCALE_IDIGITS LOCALE_ILZERO LOCALE_INEGNUMBER
		LOCALE_SNATIVEDIGITS LOCALE_SCURRENCY LOCALE_SINTLSYMBOL
		LOCALE_SMONDECIMALSEP LOCALE_SMONTHOUSANDSEP LOCALE_SMONGROUPING
		LOCALE_ICURRDIGITS LOCALE_IINTLCURRDIGITS LOCALE_ICURRENCY
		LOCALE_INEGCURR LOCALE_SDATE LOCALE_STIME LOCALE_SSHORTDATE
		LOCALE_SLONGDATE LOCALE_STIMEFORMAT LOCALE_IDATE LOCALE_ILDATE
		LOCALE_ITIME LOCALE_ITIMEMARKPOSN LOCALE_ICENTURY LOCALE_ITLZERO
		LOCALE_IDAYLZERO LOCALE_IMONLZERO LOCALE_S1159 LOCALE_S2359
		LOCALE_ICALENDARTYPE LOCALE_IOPTIONALCALENDAR
		LOCALE_IFIRSTDAYOFWEEK LOCALE_IFIRSTWEEKOFYEAR LOCALE_SDAYNAME1
		LOCALE_SDAYNAME2 LOCALE_SDAYNAME3 LOCALE_SDAYNAME4
		LOCALE_SDAYNAME5 LOCALE_SDAYNAME6 LOCALE_SDAYNAME7
		LOCALE_SABBREVDAYNAME1 LOCALE_SABBREVDAYNAME2
		LOCALE_SABBREVDAYNAME3 LOCALE_SABBREVDAYNAME4
		LOCALE_SABBREVDAYNAME5 LOCALE_SABBREVDAYNAME6
		LOCALE_SABBREVDAYNAME7 LOCALE_SMONTHNAME1 LOCALE_SMONTHNAME2
		LOCALE_SMONTHNAME3 LOCALE_SMONTHNAME4 LOCALE_SMONTHNAME5
		LOCALE_SMONTHNAME6 LOCALE_SMONTHNAME7 LOCALE_SMONTHNAME8
		LOCALE_SMONTHNAME9 LOCALE_SMONTHNAME10 LOCALE_SMONTHNAME11
		LOCALE_SMONTHNAME12 LOCALE_SMONTHNAME13 LOCALE_SABBREVMONTHNAME1
		LOCALE_SABBREVMONTHNAME2 LOCALE_SABBREVMONTHNAME3
		LOCALE_SABBREVMONTHNAME4 LOCALE_SABBREVMONTHNAME5
		LOCALE_SABBREVMONTHNAME6 LOCALE_SABBREVMONTHNAME7
		LOCALE_SABBREVMONTHNAME8 LOCALE_SABBREVMONTHNAME9
		LOCALE_SABBREVMONTHNAME10 LOCALE_SABBREVMONTHNAME11
		LOCALE_SABBREVMONTHNAME12 LOCALE_SABBREVMONTHNAME13
		LOCALE_SPOSITIVESIGN LOCALE_SNEGATIVESIGN LOCALE_IPOSSIGNPOSN
		LOCALE_INEGSIGNPOSN LOCALE_IPOSSYMPRECEDES LOCALE_IPOSSEPBYSPACE
		LOCALE_INEGSYMPRECEDES LOCALE_INEGSEPBYSPACE)]
);

foreach my $tag (keys %EXPORT_TAGS) {
    push @EXPORT_OK, @{$EXPORT_TAGS{$tag}};
}

# Character Type Flags
sub CT_CTYPE1		   { 0x0001; }	# ctype 1 information
sub CT_CTYPE2		   { 0x0002; }	# ctype 2 information
sub CT_CTYPE3		   { 0x0004; }	# ctype 3 information

# Character Type 1 Bits
sub C1_UPPER		   { 0x0001; }	# upper case
sub C1_LOWER		   { 0x0002; }	# lower case
sub C1_DIGIT		   { 0x0004; }	# decimal digits
sub C1_SPACE		   { 0x0008; }	# spacing characters
sub C1_PUNCT		   { 0x0010; }	# punctuation characters
sub C1_CNTRL		   { 0x0020; }	# control characters
sub C1_BLANK		   { 0x0040; }	# blank characters
sub C1_XDIGIT		   { 0x0080; }	# other digits
sub C1_ALPHA		   { 0x0100; }	# any letter

# Character Type 2 Bits
sub C2_LEFTTORIGHT	   { 0x1; }	# left to right
sub C2_RIGHTTOLEFT	   { 0x2; }	# right to left
sub C2_EUROPENUMBER	   { 0x3; }	# European number, digit
sub C2_EUROPESEPARATOR	   { 0x4; }	# European numeric separator
sub C2_EUROPETERMINATOR	   { 0x5; }	# European numeric terminator
sub C2_ARABICNUMBER	   { 0x6; }	# Arabic number
sub C2_COMMONSEPARATOR	   { 0x7; }	# common numeric separator
sub C2_BLOCKSEPARATOR	   { 0x8; }	# block separator
sub C2_SEGMENTSEPARATOR	   { 0x9; }	# segment separator
sub C2_WHITESPACE	   { 0xA; }	# white space
sub C2_OTHERNEUTRAL	   { 0xB; }	# other neutrals
sub C2_NOTAPPLICABLE	   { 0x0; }	# no implicit directionality

# Character Type 3 Bits
sub C3_NONSPACING	   { 0x0001; }	# nonspacing character
sub C3_DIACRITIC	   { 0x0002; }	# diacritic mark
sub C3_VOWELMARK	   { 0x0004; }	# vowel mark
sub C3_SYMBOL		   { 0x0008; }	# symbols
sub C3_KATAKANA		   { 0x0010; }
sub C3_HIRAGANA		   { 0x0020; }
sub C3_HALFWIDTH	   { 0x0040; }
sub C3_FULLWIDTH	   { 0x0080; }
sub C3_IDEOGRAPH	   { 0x0100; }
sub C3_KASHIDA		   { 0x0200; }
sub C3_ALPHA		   { 0x8000; }
sub C3_NOTAPPLICABLE	   { 0x0; }	# ctype 3 is not applicable

# String Flags
sub NORM_IGNORECASE	   { 0x0001; }	# ignore case
sub NORM_IGNORENONSPACE	   { 0x0002; }	# ignore nonspacing chars
sub NORM_IGNORESYMBOLS	   { 0x0004; }	# ignore symbols
sub NORM_IGNOREWIDTH	   { 0x0008; }	# ignore width
sub NORM_IGNOREKANATYPE	   { 0x0040; }	# ignore kanatype
sub NORM_IGNOREKASHIDA	   { 0x40000;}	# ignore Arabic kashida chars

# Locale Dependent Mapping Flags
sub LCMAP_LOWERCASE	   { 0x0100; }	# lower case letters
sub LCMAP_UPPERCASE	   { 0x0200; }	# upper case letters
sub LCMAP_SORTKEY	   { 0x0400; }	# WC sort key (normalize)
sub LCMAP_HALFWIDTH	   { 0x0800; }	# narrow pitch case letters
sub LCMAP_FULLWIDTH	   { 0x1000; }	# wide picth case letters
sub LCMAP_HIRAGANA	   { 0x2000; }	# map katakana to hiragana
sub LCMAP_KATAKANA	   { 0x4000; }	# map hiragana to katakana

# Primary Language Identifier
sub LANG_NEUTRAL	   { 0x00; }
sub LANG_ALBANIAN	   { 0x1c; }
sub LANG_ARABIC		   { 0x01; }
sub LANG_BAHASA		   { 0x21; }
sub LANG_BULGARIAN	   { 0x02; }
sub LANG_CATALAN	   { 0x03; }
sub LANG_CHINESE	   { 0x04; }
sub LANG_CZECH		   { 0x05; }
sub LANG_DANISH		   { 0x06; }
sub LANG_DUTCH		   { 0x13; }
sub LANG_ENGLISH	   { 0x09; }
sub LANG_FINNISH	   { 0x0b; }
sub LANG_FRENCH		   { 0x0c; }
sub LANG_GERMAN		   { 0x07; }
sub LANG_GREEK		   { 0x08; }
sub LANG_HEBREW		   { 0x0d; }
sub LANG_HUNGARIAN	   { 0x0e; }
sub LANG_ICELANDIC	   { 0x0f; }
sub LANG_ITALIAN	   { 0x10; }
sub LANG_JAPANESE	   { 0x11; }
sub LANG_KOREAN		   { 0x12; }
sub LANG_NORWEGIAN	   { 0x14; }
sub LANG_POLISH		   { 0x15; }
sub LANG_PORTUGUESE	   { 0x16; }
sub LANG_RHAETO_ROMAN	   { 0x17; }
sub LANG_ROMANIAN	   { 0x18; }
sub LANG_RUSSIAN	   { 0x19; }
sub LANG_SERBO_CROATIAN	   { 0x1a; }
sub LANG_SLOVAK		   { 0x1b; }
sub LANG_SPANISH	   { 0x0a; }
sub LANG_SWEDISH	   { 0x1d; }
sub LANG_THAI		   { 0x1e; }
sub LANG_TURKISH	   { 0x1f; }
sub LANG_URDU		   { 0x20; }

# Sublanguage Identifier
sub SUBLANG_NEUTRAL		    { 0x00; } # language neutral
sub SUBLANG_DEFAULT		    { 0x01; } # user default
sub SUBLANG_SYS_DEFAULT		    { 0x02; } # system default
sub SUBLANG_CHINESE_SIMPLIFIED	    { 0x02; } # Chinese (Simplified)
sub SUBLANG_CHINESE_TRADITIONAL	    { 0x01; } # Chinese (Traditional)
sub SUBLANG_DUTCH		    { 0x01; } # Dutch
sub SUBLANG_DUTCH_BELGIAN	    { 0x02; } # Dutch (Belgian)
sub SUBLANG_ENGLISH_US		    { 0x01; } # English (USA)
sub SUBLANG_ENGLISH_UK		    { 0x02; } # English (UK)
sub SUBLANG_ENGLISH_AUS		    { 0x03; } # English (Australian)
sub SUBLANG_ENGLISH_CAN		    { 0x04; } # English (Canadian)
sub SUBLANG_ENGLISH_NZ		    { 0x05; } # English (New Zealand)
sub SUBLANG_ENGLISH_EIRE	    { 0x06; } # English (Irish)
sub SUBLANG_FRENCH		    { 0x01; } # French
sub SUBLANG_FRENCH_BELGIAN	    { 0x02; } # French (Belgian)
sub SUBLANG_FRENCH_CANADIAN	    { 0x03; } # French (Canadian)
sub SUBLANG_FRENCH_SWISS	    { 0x04; } # French (Swiss)
sub SUBLANG_GERMAN		    { 0x01; } # German
sub SUBLANG_GERMAN_SWISS	    { 0x02; } # German (Swiss)
sub SUBLANG_GERMAN_AUSTRIAN	    { 0x03; } # German (Austrian)
sub SUBLANG_ITALIAN		    { 0x01; } # Italian
sub SUBLANG_ITALIAN_SWISS	    { 0x02; } # Italian (Swiss)
sub SUBLANG_NORWEGIAN_BOKMAL	    { 0x01; } # Norwegian (Bokmal)
sub SUBLANG_NORWEGIAN_NYNORSK	    { 0x02; } # Norwegian (Nynorsk)
sub SUBLANG_PORTUGUESE		    { 0x02; } # Portuguese
sub SUBLANG_PORTUGUESE_BRAZILIAN    { 0x01; } # Portuguese (Brazilian)
sub SUBLANG_SERBO_CROATIAN_CYRILLIC { 0x02; } # Serbo-Croatian (Cyrillic)
sub SUBLANG_SERBO_CROATIAN_LATIN    { 0x01; } # Croato-Serbian (Latin)
sub SUBLANG_SPANISH		    { 0x01; } # Spanish
sub SUBLANG_SPANISH_MEXICAN	    { 0x02; } # Spanish (Mexican)
sub SUBLANG_SPANISH_MODERN	    { 0x03; } # Spanish (Modern)

# Country codes
sub CTRY_DEFAULT	      { 0;	}
sub CTRY_AUSTRALIA	      { 61;	} # Australia
sub CTRY_AUSTRIA	      { 43;	} # Austria
sub CTRY_BELGIUM	      { 32;	} # Belgium
sub CTRY_BRAZIL		      { 55;	} # Brazil
sub CTRY_CANADA		      { 2;	} # Canada
sub CTRY_DENMARK	      { 45;	} # Denmark
sub CTRY_FINLAND	      { 358;	} # Finland
sub CTRY_FRANCE		      { 33;	} # France
sub CTRY_GERMANY	      { 49;	} # Germany
sub CTRY_ICELAND	      { 354;	} # Iceland
sub CTRY_IRELAND	      { 353;	} # Ireland
sub CTRY_ITALY		      { 39;	} # Italy
sub CTRY_JAPAN		      { 81;	} # Japan
sub CTRY_MEXICO		      { 52;	} # Mexico
sub CTRY_NETHERLANDS	      { 31;	} # Netherlands
sub CTRY_NEW_ZEALAND	      { 64;	} # New Zealand
sub CTRY_NORWAY		      { 47;	} # Norway
sub CTRY_PORTUGAL	      { 351;	} # Portugal
sub CTRY_PRCHINA	      { 86;	} # PR China
sub CTRY_SOUTH_KOREA	      { 82;	} # South Korea
sub CTRY_SPAIN		      { 34;	} # Spain
sub CTRY_SWEDEN		      { 46;	} # Sweden
sub CTRY_SWITZERLAND	      { 41;	} # Switzerland
sub CTRY_TAIWAN		      { 886;	} # Taiwan
sub CTRY_UNITED_KINGDOM	      { 44;	} # United Kingdom
sub CTRY_UNITED_STATES	      { 1;	} # United States

# Locale Types
sub LOCALE_NOUSEROVERRIDE { 0x80000000; } # OR in to avoid user override
sub LOCALE_ILANGUAGE	      { 0x0001; } # language id
sub LOCALE_SLANGUAGE	      { 0x0002; } # localized name of language
sub LOCALE_SENGLANGUAGE	      { 0x1001; } # English name of language
sub LOCALE_SABBREVLANGNAME    { 0x0003; } # abbreviated language name
sub LOCALE_SNATIVELANGNAME    { 0x0004; } # native name of language
sub LOCALE_ICOUNTRY	      { 0x0005; } # country code
sub LOCALE_SCOUNTRY	      { 0x0006; } # localized name of country
sub LOCALE_SENGCOUNTRY	      { 0x1002; } # English name of country
sub LOCALE_SABBREVCTRYNAME    { 0x0007; } # abbreviated country name
sub LOCALE_SNATIVECTRYNAME    { 0x0008; } # native name of country
sub LOCALE_IDEFAULTLANGUAGE   { 0x0009; } # default language id
sub LOCALE_IDEFAULTCOUNTRY    { 0x000A; } # default country code
sub LOCALE_IDEFAULTCODEPAGE   { 0x000B; } # default oem code page
sub LOCALE_IDEFAULTANSICODEPAGE{0x1004; } # default ansi code page
sub LOCALE_SLIST	      { 0x000C; } # list item separator
sub LOCALE_IMEASURE	      { 0x000D; } # 0 = metric, 1 = US
sub LOCALE_SDECIMAL	      { 0x000E; } # decimal separator
sub LOCALE_STHOUSAND	      { 0x000F; } # thousand separator
sub LOCALE_SGROUPING	      { 0x0010; } # digit grouping
sub LOCALE_IDIGITS	      { 0x0011; } # number of fractional digits
sub LOCALE_ILZERO	      { 0x0012; } # leading zeros for decimal
sub LOCALE_INEGNUMBER	      { 0x1010; } # negative number mode
sub LOCALE_SNATIVEDIGITS      { 0x0013; } # native ascii 0-9
sub LOCALE_SCURRENCY	      { 0x0014; } # local monetary symbol
sub LOCALE_SINTLSYMBOL	      { 0x0015; } # intl monetary symbol
sub LOCALE_SMONDECIMALSEP     { 0x0016; } # monetary decimal separator
sub LOCALE_SMONTHOUSANDSEP    { 0x0017; } # monetary thousand separator
sub LOCALE_SMONGROUPING	      { 0x0018; } # monetary grouping
sub LOCALE_ICURRDIGITS	      { 0x0019; } # # local monetary digits
sub LOCALE_IINTLCURRDIGITS    { 0x001A; } # # intl monetary digits
sub LOCALE_ICURRENCY	      { 0x001B; } # positive currency mode
sub LOCALE_INEGCURR	      { 0x001C; } # negative currency mode
sub LOCALE_SDATE	      { 0x001D; } # date separator
sub LOCALE_STIME	      { 0x001E; } # time separator
sub LOCALE_SSHORTDATE	      { 0x001F; } # short date-time separator
sub LOCALE_SLONGDATE	      { 0x0020; } # long date-time separator
sub LOCALE_STIMEFORMAT	      { 0x1003; } # time format string
sub LOCALE_IDATE	      { 0x0021; } # short date format ordering
sub LOCALE_ILDATE	      { 0x0022; } # long date format ordering
sub LOCALE_ITIME	      { 0x0023; } # time format specifier
sub LOCALE_ITIMEMARKPOSN      { 0x1005; } # time marker position
sub LOCALE_ICENTURY	      { 0x0024; } # century format specifier
sub LOCALE_ITLZERO	      { 0x0025; } # leading zeros in time field
sub LOCALE_IDAYLZERO	      { 0x0026; } # leading zeros in day field
sub LOCALE_IMONLZERO	      { 0x0027; } # leading zeros in month field
sub LOCALE_S1159	      { 0x0028; } # AM designator
sub LOCALE_S2359	      { 0x0029; } # PM designator
sub LOCALE_ICALENDARTYPE      { 0x1009; } # type of calendar specifier
sub LOCALE_IOPTIONALCALENDAR  { 0x100B; } # additional calendar types specifier
sub LOCALE_IFIRSTDAYOFWEEK    { 0x100C; } # first day of week specifier
sub LOCALE_IFIRSTWEEKOFYEAR   { 0x100D; } # first week of year specifier
sub LOCALE_SDAYNAME1	      { 0x002A; } # long name for Monday
sub LOCALE_SDAYNAME2	      { 0x002B; } # long name for Tuesday
sub LOCALE_SDAYNAME3	      { 0x002C; } # long name for Wednesday
sub LOCALE_SDAYNAME4	      { 0x002D; } # long name for Thursday
sub LOCALE_SDAYNAME5	      { 0x002E; } # long name for Friday
sub LOCALE_SDAYNAME6	      { 0x002F; } # long name for Saturday
sub LOCALE_SDAYNAME7	      { 0x0030; } # long name for Sunday
sub LOCALE_SABBREVDAYNAME1    { 0x0031; } # abbreviated name for Monday
sub LOCALE_SABBREVDAYNAME2    { 0x0032; } # abbreviated name for Tuesday
sub LOCALE_SABBREVDAYNAME3    { 0x0033; } # abbreviated name for Wednesday
sub LOCALE_SABBREVDAYNAME4    { 0x0034; } # abbreviated name for Thursday
sub LOCALE_SABBREVDAYNAME5    { 0x0035; } # abbreviated name for Friday
sub LOCALE_SABBREVDAYNAME6    { 0x0036; } # abbreviated name for Saturday
sub LOCALE_SABBREVDAYNAME7    { 0x0037; } # abbreviated name for Sunday
sub LOCALE_SMONTHNAME1	      { 0x0038; } # long name for January
sub LOCALE_SMONTHNAME2	      { 0x0039; } # long name for February
sub LOCALE_SMONTHNAME3	      { 0x003A; } # long name for March
sub LOCALE_SMONTHNAME4	      { 0x003B; } # long name for April
sub LOCALE_SMONTHNAME5	      { 0x003C; } # long name for May
sub LOCALE_SMONTHNAME6	      { 0x003D; } # long name for June
sub LOCALE_SMONTHNAME7	      { 0x003E; } # long name for July
sub LOCALE_SMONTHNAME8	      { 0x003F; } # long name for August
sub LOCALE_SMONTHNAME9	      { 0x0040; } # long name for September
sub LOCALE_SMONTHNAME10	      { 0x0041; } # long name for October
sub LOCALE_SMONTHNAME11	      { 0x0042; } # long name for November
sub LOCALE_SMONTHNAME12	      { 0x0043; } # long name for December
sub LOCALE_SMONTHNAME13	      { 0x100E; } # long name for 13th month
sub LOCALE_SABBREVMONTHNAME1  { 0x0044; } # abbreviated name for January
sub LOCALE_SABBREVMONTHNAME2  { 0x0045; } # abbreviated name for February
sub LOCALE_SABBREVMONTHNAME3  { 0x0046; } # abbreviated name for March
sub LOCALE_SABBREVMONTHNAME4  { 0x0047; } # abbreviated name for April
sub LOCALE_SABBREVMONTHNAME5  { 0x0048; } # abbreviated name for May
sub LOCALE_SABBREVMONTHNAME6  { 0x0049; } # abbreviated name for June
sub LOCALE_SABBREVMONTHNAME7  { 0x004A; } # abbreviated name for July
sub LOCALE_SABBREVMONTHNAME8  { 0x004B; } # abbreviated name for August
sub LOCALE_SABBREVMONTHNAME9  { 0x004C; } # abbreviated name for September
sub LOCALE_SABBREVMONTHNAME10 { 0x004D; } # abbreviated name for October
sub LOCALE_SABBREVMONTHNAME11 { 0x004E; } # abbreviated name for November
sub LOCALE_SABBREVMONTHNAME12 { 0x004F; } # abbreviated name for December
sub LOCALE_SABBREVMONTHNAME13 { 0x100F; } # abbreviated name for 13th month
sub LOCALE_SPOSITIVESIGN      { 0x0050; } # positive sign
sub LOCALE_SNEGATIVESIGN      { 0x0051; } # negative sign
sub LOCALE_IPOSSIGNPOSN	      { 0x0052; } # positive sign position
sub LOCALE_INEGSIGNPOSN	      { 0x0053; } # negative sign position
sub LOCALE_IPOSSYMPRECEDES    { 0x0054; } # mon sym precedes pos amt
sub LOCALE_IPOSSEPBYSPACE     { 0x0055; } # mon sym sep by space from pos
sub LOCALE_INEGSYMPRECEDES    { 0x0056; } # mon sym precedes neg amt
sub LOCALE_INEGSEPBYSPACE     { 0x0057; } # mon sym sep by space from neg

# Language Identifier Creation/Extraction Functions
sub MAKELANGID	   { my ($p,$s) = @_; (($s & 0xffff) << 10) | ($p & 0xffff); }
sub PRIMARYLANGID  { my $lgid = shift; $lgid & 0x3ff; }
sub SUBLANGID	   { my $lgid = shift; ($lgid >> 10) & 0x3f; }

# Locale Identifier Creation Functions
sub MAKELCID	   { my $lgid = shift; $lgid & 0xffff; }
sub LANGIDFROMLCID { my $lcid = shift; $lcid & 0xffff; }

1;

__END__

=head1 NAME

Win32::OLE::NLS - OLE National Language Support

=head1 SYNOPSIS

	missing

=head1 DESCRIPTION

This module provides access to the national language features in
the OLENLS.DLL. It is still VERY incomplete.

=head2 Functions/Methods

=over 8

=item GetLocaleInfo(LCID,LCTYPE)

Retrieve locale setting LCTYPE from locale specified by id LCID.

=back

=head1 AUTHORS/COPYRIGHT

This module is part of the Win32::OLE distribution.

=cut
