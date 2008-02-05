/* OLE.xs
 *
 *  (c) 1995 Microsoft Corporation. All rights reserved.
 *  Developed by ActiveWare Internet Corp., now known as
 *  ActiveState Tool Corp., http://www.ActiveState.com
 *
 *  Other modifications Copyright (c) 1997, 1998 by Gurusamy Sarathy
 *  <gsar@umich.edu> and Jan Dubois <jan.dubois@ibm.net>
 *
 *  You may distribute under the terms of either the GNU General Public
 *  License or the Artistic License, as specified in the README file.
 *
 *
 * File contents:
 *
 * - C helper routines
 * - Package Win32::OLE             Constructor and method invocation
 * - Package Win32::OLE::Tie        Implements properties as tied hash
 * - Package Win32::OLE::Const      Load application constants from type library
 * - Package Win32::OLE::Enum       OLE collection enumeration
 * - Package Win32::OLE::Variant    Implements Perl VARIANT objects
 * - Package Win32::OLE::NLS        National Language Support
 * - Package Win32::OLE::TypeLib    Type library access
 * - Package Win32::OLE::TypeInfo   Type info access
 *
 */

#define MY_VERSION "Win32::OLE(" XS_VERSION ")"

#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#define _WIN32_DCOM
#include <windows.h>

#ifdef _DEBUG
#   include <crtdbg.h>
#   define DEBUGBREAK _CrtDbgBreak()
#else
#   define DEBUGBREAK
#endif

#if defined (__cplusplus)
extern "C" {
#endif

#define MIN_PERL_DEFINE
#define NO_XSLOCKS
#include "EXTERN.h"
#include "perl.h"
#include "XSub.h"
#include "patchlevel.h"

#if (PATCHLEVEL < 4) || ((PATCHLEVEL == 4) && (SUBVERSION < 1))
#   error Win32::OLE module requires Perl 5.004_01 or later
#endif

#if (PATCHLEVEL < 5)
#   ifndef PL_dowarn
#	define PL_dowarn	dowarn
#	define PL_sv_undef	sv_undef
#	define PL_sv_yes	sv_yes
#	define PL_sv_no		sv_no
#   endif
#   define PL_hints		hints
#   define PL_modglobal		modglobal
#endif

#ifndef CPERLarg
#   define CPERLarg
#   define CPERLarg_
#   define PERL_OBJECT_THIS
#   define PERL_OBJECT_THIS_
#endif

#if !defined(_DEBUG)
#   define DBG(a)
#else
#   define DBG(a)  MyDebug a
void
MyDebug(const char *pat, ...)
{
    va_list args;
    va_start(args, pat);

#if 1
    char szBuffer[512];
    vsprintf(szBuffer, pat, args);
    OutputDebugString(szBuffer);
#else
    PerlIO_vprintf(PerlIO_stderr(), pat, args);
    PerlIO_flush(PerlIO_stderr());
#endif

    va_end(args);
}
#endif

/* constants */
static const DWORD WINOLE_MAGIC = 0x12344321;
static const DWORD WINOLEENUM_MAGIC = 0x12344322;
static const DWORD WINOLEVARIANT_MAGIC = 0x12344323;
static const DWORD WINOLETYPELIB_MAGIC = 0x12344324;
static const DWORD WINOLETYPEINFO_MAGIC = 0x12344325;

static const LCID lcidSystemDefault = 2 << 10;
/* static const LCID lcidDefault = 0; language neutral */
static const LCID lcidDefault = lcidSystemDefault;
static const UINT cpDefault = CP_ACP;
static char PERL_OLE_ID[] = "___Perl___OleObject___";
static const int PERL_OLE_IDLEN = sizeof(PERL_OLE_ID)-1;

static const int OLE_BUF_SIZ = 256;

/* class names */
static char szWINOLE[] = "Win32::OLE";
static char szWINOLEENUM[] = "Win32::OLE::Enum";
static char szWINOLEVARIANT[] = "Win32::OLE::Variant";
static char szWINOLETIE[] = "Win32::OLE::Tie";
static char szWINOLETYPELIB[] = "Win32::OLE::TypeLib";
static char szWINOLETYPEINFO[] = "Win32::OLE::TypeInfo";

/* class variable names */
static char LCID_NAME[] = "LCID";
static const int LCID_LEN = sizeof(LCID_NAME)-1;
static char CP_NAME[] = "CP";
static const int CP_LEN = sizeof(CP_NAME)-1;
static char WARN_NAME[] = "Warn";
static const int WARN_LEN = sizeof(WARN_NAME)-1;
static char LASTERR_NAME[] = "LastError";
static const int LASTERR_LEN = sizeof(LASTERR_NAME)-1;
static char TIE_NAME[] = "Tie";
static const int TIE_LEN = sizeof(TIE_NAME)-1;

#define COINIT_OLEINITIALIZE -1

typedef HRESULT (STDAPICALLTYPE FNCOINITIALIZEEX)(LPVOID, DWORD);
typedef void (STDAPICALLTYPE FNCOUNINITIALIZE)(void);
typedef HRESULT (STDAPICALLTYPE FNCOCREATEINSTANCEEX)
    (REFCLSID, IUnknown*, DWORD, COSERVERINFO*, DWORD, MULTI_QI*);

typedef struct _tagOBJECTHEADER OBJECTHEADER;

/* per interpreter variables */
typedef struct
{
    CRITICAL_SECTION CriticalSection;
    OBJECTHEADER *pObj;
    BOOL bInitialized;

    /* DCOM function addresses are resolved dynamically */
    HINSTANCE hOLE32;
    FNCOINITIALIZEEX     *pfnCoInitializeEx;
    FNCOUNINITIALIZE     *pfnCoUninitialize;
    FNCOCREATEINSTANCEEX *pfnCoCreateInstanceEx;

}   PERINTERP;

#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
#   if (PATCHLEVEL == 4) && (SUBVERSION < 68)
#       define dPERINTERP                                                 \
           SV *interp = perl_get_sv(MY_VERSION, FALSE);                   \
           if (interp == NULL || !SvIOK(interp))                          \
               warn(MY_VERSION ": Per-interpreter data not initialized"); \
           PERINTERP *pInterp = (PERINTERP*)SvIV(interp)
#   else
#	define dPERINTERP                                                 \
           SV **pinterp = hv_fetch(PL_modglobal, MY_VERSION,              \
                                   sizeof(MY_VERSION)-1, FALSE);          \
           if (pinterp == NULL || *pinterp == NULL || !SvIOK(*pinterp))   \
               warn(MY_VERSION ": Per-interpreter data not initialized"); \
	   PERINTERP *pInterp = (PERINTERP*)SvIV(*pinterp)
#   endif
#   define INTERP pInterp
#else
static PERINTERP Interp;
#   define dPERINTERP extern int errno
#   define INTERP (&Interp)
#endif

#define g_pObj            (INTERP->pObj)
#define g_bInitialized    (INTERP->bInitialized)
#define g_CriticalSection (INTERP->CriticalSection)

#define g_hOLE32                (INTERP->hOLE32)
#define g_pfnCoInitializeEx     (INTERP->pfnCoInitializeEx)
#define g_pfnCoUninitialize     (INTERP->pfnCoUninitialize)
#define g_pfnCoCreateInstanceEx (INTERP->pfnCoCreateInstanceEx)

/* common object header */
typedef struct _tagOBJECTHEADER
{
    long lMagic;
    OBJECTHEADER *pNext;
    OBJECTHEADER *pPrevious;
#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    PERINTERP    *pInterp;
#endif
}   OBJECTHEADER;

/* Win32::OLE object */
typedef struct
{
    OBJECTHEADER header;

    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo;
    IEnumVARIANT *pEnum;

    HV *self;
    HV *hashTable;
    SV *destroy;

    unsigned short cFuncs;
    unsigned short cVars;
    unsigned int   PropIndex;

}   WINOLEOBJECT;

/* Win32::OLE::Enum object */
typedef struct
{
    OBJECTHEADER header;

    IEnumVARIANT *pEnum;

}   WINOLEENUMOBJECT;

/* Win32::OLE::Variant object */
typedef struct
{
    OBJECTHEADER header;

    VARIANT variant;
    VARIANT byref;

}   WINOLEVARIANTOBJECT;

/* Win32::OLE::TypeLib object */
typedef struct
{
    OBJECTHEADER header;

    ITypeLib  *pTypeLib;
    TLIBATTR  *pTLibAttr;

}   WINOLETYPELIBOBJECT;

/* Win32::OLE::TypeInfo object */
typedef struct
{
    OBJECTHEADER header;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

}   WINOLETYPEINFOOBJECT;

/* forward declarations */
HRESULT SetSVFromVariantEx(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash,
			   BOOL bByRefObj=FALSE);

/* The following function from IO.xs is in the core starting with 5.004_63 */
#if (PATCHLEVEL == 4) && (SUBVERSION < 63)
void
newCONSTSUB(HV *stash, char *name, SV *sv)
{
#ifdef dTHR
    dTHR;
#endif
    U32 oldhints = PL_hints;
    HV *old_cop_stash = curcop->cop_stash;
    HV *old_curstash = curstash;
    line_t oldline = curcop->cop_line;
    curcop->cop_line = copline;

    PL_hints &= ~HINT_BLOCK_SCOPE;
    if(stash)
	curstash = curcop->cop_stash = stash;

    newSUB(
	start_subparse(FALSE, 0),
	newSVOP(OP_CONST, 0, newSVpv(name,0)),
	newSVOP(OP_CONST, 0, &sv_no),	/* SvPV(&sv_no) == "" -- GMB */
	newSTATEOP(0, Nullch, newSVOP(OP_CONST, 0, sv))
    );

    PL_hints = oldhints;
    curcop->cop_stash = old_cop_stash;
    curstash = old_curstash;
    curcop->cop_line = oldline;
}
#endif

BOOL
IsLocalMachine(CPERLarg_ char *pszMachine)
{
    char szComputerName[MAX_COMPUTERNAME_LENGTH+1];
    DWORD dwSize = sizeof(szComputerName);
    char *pszName = pszMachine;

    while (*pszName == '\\')
	++pszName;

    if (*pszName == '\0')
	return TRUE;

    /* Check against local computer name (from registry) */
    if (GetComputerName(szComputerName, &dwSize)
	&& stricmp(pszName, szComputerName) == 0)
	return TRUE;

    /* gethostname(), gethostbyname() and inet_addr() all call proxy functions
     * in the Perl socket layer wrapper in win32sck.c. Therefore calling
     * WSAStartup() here is not necessary.
     */

    /* Determine main host name of local machine */
    char szBuffer[200];
    if (gethostname(szBuffer, sizeof(szBuffer)) != 0)
	return FALSE;

    /* Copy list of addresses for local machine */
    struct hostent *pHostEnt = gethostbyname(szBuffer);
    if (pHostEnt == NULL)
	return FALSE;

    if (pHostEnt->h_addrtype != PF_INET || pHostEnt->h_length != 4) {
	warn(MY_VERSION ": IsLocalMachine() gethostbyname failure");
	return FALSE;
    }

    int index;
    int count = 0;
    char *pLocal;
    while (pHostEnt->h_addr_list[count] != NULL)
	++count;

    New(0, pLocal, 4*count, char);
    for (index = 0 ; index < count ; ++index)
	memcpy(pLocal+4*index, pHostEnt->h_addr_list[index], 4);

    /* Determine addresses of remote machine */
    unsigned long ulRemoteAddr;
    char *pRemote[2] = {NULL, NULL};
    char **ppRemote = &pRemote[0];

    if (isdigit(*pszMachine)) {
	/* Convert numeric dotted address */
	ulRemoteAddr = inet_addr(pszMachine);
	if (ulRemoteAddr != INADDR_NONE)
	    pRemote[0] = (char*)&ulRemoteAddr;
    }
    else {
	/* Lookup addresses for remote host name */
	pHostEnt = gethostbyname(pszMachine);
	if (pHostEnt != NULL)
	    if (pHostEnt->h_addrtype == PF_INET && pHostEnt->h_length == 4)
		ppRemote = pHostEnt->h_addr_list;
    }

    /* Compare list of addresses of remote machine against local addresses */
    while (*ppRemote != NULL) {
	for (index = 0 ; index < count ; ++index)
	    if (memcmp(pLocal+4*index, *ppRemote, 4) == 0) {
		Safefree(pLocal);
		return TRUE;
	    }
	++ppRemote;
    }

    Safefree(pLocal);
    return FALSE;

}   /* IsLocalMachine */

HRESULT
CLSIDFromRemoteRegistry(CPERLarg_ char *pszHost, char *pszProgID, CLSID *pCLSID)
{
    HKEY hKeyLocalMachine;
    HKEY hKeyProgID;
    LONG err;
    HRESULT hr = S_OK;
    STRLEN len;

    err = RegConnectRegistry(pszHost, HKEY_LOCAL_MACHINE, &hKeyLocalMachine);
    if (err != ERROR_SUCCESS)
	return HRESULT_FROM_WIN32(err);

    SV *subkey = sv_2mortal(newSVpv("SOFTWARE\\Classes\\", 0));
    sv_catpv(subkey, pszProgID);
    sv_catpv(subkey, "\\CLSID");

    err = RegOpenKeyEx(hKeyLocalMachine, SvPV(subkey, len), 0, KEY_READ,
		       &hKeyProgID);
    if (err != ERROR_SUCCESS)
	hr = HRESULT_FROM_WIN32(err);
    else {
	DWORD dwType;
	char szCLSID[100];
	DWORD dwLength = sizeof(szCLSID);

	err = RegQueryValueEx(hKeyProgID, "", NULL, &dwType,
			      (unsigned char*)&szCLSID, &dwLength);
	if (err != ERROR_SUCCESS)
	    hr = HRESULT_FROM_WIN32(err);
	else if (dwType == REG_SZ) {
	    OLECHAR wszCLSID[sizeof(szCLSID)];

	    MultiByteToWideChar(CP_ACP, 0, szCLSID, -1,
				wszCLSID, sizeof(szCLSID));
	    hr = CLSIDFromString(wszCLSID, pCLSID);
	}
	else /* XXX maybe there is a more appropriate error code? */
	    hr = HRESULT_FROM_WIN32(ERROR_CANTREAD);

	RegCloseKey(hKeyProgID);
    }

    RegCloseKey(hKeyLocalMachine);
    return hr;

}   /* CLSIDFromRemoteRegistry */

/* The following strategy is used to avoid the limitations of hardcoded
 * buffer sizes: Conversion between wide char and multibyte strings
 * is performed by GetMultiByte and GetWideChar respectively. The
 * caller passes a default buffer and size. If the buffer is too small
 * then the conversion routine allocates a new buffer that is big enough.
 * The caller must free this buffer using the ReleaseBuffer function. */

inline void
ReleaseBuffer(CPERLarg_ void *pszHeap, void *pszStack)
{
    if (pszHeap != pszStack && pszHeap != NULL)
	Safefree(pszHeap);
}

char *
GetMultiByte(CPERLarg_ OLECHAR *wide, char *psz, int len, UINT cp)
{
    int count;

    if (psz != NULL) {
	if (wide == NULL) {
	    *psz = (char) 0;
	    return psz;
	}
	count = WideCharToMultiByte(cp, 0, wide, -1, psz, len, NULL, NULL);
	if (count > 0)
	    return psz;
    }
    else if (wide == NULL) {
	Newz(0, psz, 1, char);
	return psz;
    }

    count = WideCharToMultiByte(cp, 0, wide, -1, NULL, 0, NULL, NULL);
    if (count == 0) { /* should never happen */
	warn(MY_VERSION ": GetMultiByte() failure: %lu", GetLastError());
	DEBUGBREAK;
	if (psz == NULL)
	    New(0, psz, 1, char);
	*psz = (char) 0;
	return psz;
    }

    Newz(0, psz, count, char);
    WideCharToMultiByte(cp, 0, wide, -1, psz, count, NULL, NULL);
    return psz;

}   /* GetMultiByte */

SV *
sv_setwide(CPERLarg_ SV *sv, OLECHAR *wide, UINT cp)
{
    char szBuffer[OLE_BUF_SIZ];
    char *pszBuffer;

    pszBuffer = GetMultiByte(PERL_OBJECT_THIS_ wide,
			     szBuffer, sizeof(szBuffer), cp);
    if (sv == NULL)
	sv = newSVpv(pszBuffer, 0);
    else
	sv_setpv(sv, pszBuffer);
    ReleaseBuffer(PERL_OBJECT_THIS_ pszBuffer, szBuffer);
    return sv;
}

OLECHAR *
GetWideChar(CPERLarg_ char *psz, OLECHAR *wide, int len, UINT cp)
{
    /* Note: len is number of OLECHARs, not bytes! */
    int count;

    if (wide != NULL) {
	if (psz == NULL) {
	    *wide = (OLECHAR) 0;
	    return wide;
	}
	count = MultiByteToWideChar(cp, 0, psz, -1, wide, len);
	if (count > 0)
	    return wide;
    }
    else if (psz == NULL) {
	Newz(0, wide, 1, OLECHAR);
	return wide;
    }

    count = MultiByteToWideChar(cp, 0, psz, -1, NULL, 0);
    if (count == 0) {
	warn(MY_VERSION ": GetWideChar() failure: %lu", GetLastError());
	DEBUGBREAK;
	if (wide == NULL)
	    New(0, wide, 1, OLECHAR);
	*wide = (OLECHAR) 0;
	return wide;
    }

    Newz(0, wide, count, OLECHAR);
    MultiByteToWideChar(cp, 0, psz, -1, wide, count);
    return wide;

}   /* GetWideChar */

HV *
GetStash(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv))
	return SvSTASH(SvRV(sv));
    else if (SvPOK(sv))
	return gv_stashsv(sv, TRUE);
    else
	return (HV *)&PL_sv_undef;

}   /* GetStash */

HV *
GetWin32OleStash(CPERLarg_ SV *sv)
{
    SV *pkg;
    STRLEN len;

    if (sv_isobject(sv))
	pkg = newSVpv(HvNAME(SvSTASH(SvRV(sv))), 0);
    else if (SvPOK(sv))
	pkg = newSVpv(SvPV(sv, len), len);
    else
	pkg = newSVpv(szWINOLE, 0); /* should never happen */

    char *pszColon = strrchr(SvPVX(pkg), ':');
    if (pszColon != NULL) {
	--pszColon;
	while (pszColon > SvPVX(pkg) && *pszColon == ':')
	    --pszColon;
	SvCUR(pkg) = pszColon - SvPVX(pkg) + 1;
	SvPVX(pkg)[SvCUR(pkg)] = '\0';
    }

    HV *stash = gv_stashsv(pkg, TRUE);
    SvREFCNT_dec(pkg);
    return stash;

}   /* GetWin32OleStash */

IV
QueryPkgVar(CPERLarg_ HV *stash, char *var, STRLEN len, IV def=0)
{
    SV *sv;
    GV **gv = (GV**)hv_fetch(stash, var, len, FALSE);

    if (gv != NULL && (sv = GvSV(*gv)) != NULL && SvIOK(sv)) {
	DBG(("QueryPkgVar(%s) returns %d\n", var, SvIV(sv)));
	return SvIV(sv);
    }

    return def;
}

void
SetLastOleError(CPERLarg_ HV *stash, HRESULT hr=S_OK, char *pszMsg=NULL)
{
    STRLEN len;

    /* Find $Win32::OLE::LastError */
    SV *sv = sv_2mortal(newSVpv(HvNAME(stash), 0));
    sv_catpvn(sv, "::", 2);
    sv_catpvn(sv, LASTERR_NAME, LASTERR_LEN);
    SV *lasterr = perl_get_sv(SvPV(sv, len), TRUE);
    if (lasterr == NULL) {
	warn(MY_VERSION ": SetLastOleError: couldnot create variable %s",
	     LASTERR_NAME);
	DEBUGBREAK;
	return;
    }

    sv_setiv(lasterr, (IV)hr);
    if (pszMsg != NULL) {
	sv_setpv(lasterr, pszMsg);
	SvIOK_on(lasterr);
    }
}

void
ReportOleError(CPERLarg_ HV *stash, HRESULT hr, EXCEPINFO *pExcep=NULL,
	       SV *svAdd=NULL)
{
    dSP;

    IV warnlvl = QueryPkgVar(PERL_OBJECT_THIS_ stash, WARN_NAME, WARN_LEN);
    SV *sv = sv_2mortal(newSVpv("",0));
    STRLEN len;

    /* start with exception info */
    if (pExcep != NULL && (pExcep->bstrSource != NULL ||
			   pExcep->bstrDescription != NULL )) {
	char szSource[80] = "<Unknown Source>";
	char szDesc[200] = "<No description provided>";

	char *pszSource = szSource;
	char *pszDesc = szDesc;

	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);

	if (pExcep->bstrSource != NULL)
	    pszSource = GetMultiByte(PERL_OBJECT_THIS_ pExcep->bstrSource,
				     szSource, sizeof(szSource), cp);

	if (pExcep->bstrDescription != NULL)
	    pszDesc = GetMultiByte(PERL_OBJECT_THIS_ pExcep->bstrDescription,
				   szDesc, sizeof(szDesc), cp);

	sv_setpvf(sv, "OLE exception from \"%s\":\n\n%s\n\n",
		  pszSource, pszDesc);

	ReleaseBuffer(PERL_OBJECT_THIS_ pszSource, szSource);
	ReleaseBuffer(PERL_OBJECT_THIS_ pszDesc, szDesc);
	/* SysFreeString accepts NULL too */
	SysFreeString(pExcep->bstrSource);
	SysFreeString(pExcep->bstrDescription);
	SysFreeString(pExcep->bstrHelpFile);
    }

    /* always include OLE error code */
    sv_catpvf(sv, MY_VERSION " error 0x%08x", hr);

    /* try to append ': "error text"' from message catalog */
    char *pszMsgText;
    DWORD dwCount = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
				  FORMAT_MESSAGE_FROM_SYSTEM |
				  FORMAT_MESSAGE_IGNORE_INSERTS,
				  NULL, hr, lcidSystemDefault,
				  (LPTSTR)&pszMsgText, 0, NULL);
    if (dwCount > 0) {
	sv_catpv(sv, ": \"");
	/* remove trailing dots and CRs/LFs from message */
	while (dwCount > 0 &&
	       (pszMsgText[dwCount-1] < ' ' || pszMsgText[dwCount-1] == '.'))
	    pszMsgText[--dwCount] = '\0';

	/* skip carriage returns in message text */
	char *psz = pszMsgText;
	char *pCR;
	while ((pCR = strchr(psz, '\r')) != NULL) {
	    sv_catpvn(sv, psz, pCR-psz);
	    psz = pCR+1;
	}
	if (*psz != '\0')
	    sv_catpv(sv, psz);
	sv_catpv(sv, "\"");
	LocalFree(pszMsgText);
    }

    /* add additional error details */
    if (svAdd != NULL) {
	sv_catpv(sv, "\n    ");
	sv_catsv(sv, svAdd);
    }

    /* try to keep linelength of description below 80 chars. */
    char *pLastBlank = NULL;
    char *pch = SvPV(sv, len);
    int  cch;

    for (cch = 0 ; *pch ; ++pch, ++cch) {
	if (*pch == ' ') {
	    pLastBlank = pch;
	}
	else if (*pch == '\n') {
	    pLastBlank = pch;
	    cch = 0;
	}

	if (cch > 76 && pLastBlank != NULL) {
	    *pLastBlank = '\n';
	    cch = pch - pLastBlank;
	}
    }

    SetLastOleError(PERL_OBJECT_THIS_ stash, hr, SvPV(sv, len));

    if (warnlvl > 1 || (warnlvl == 1 && PL_dowarn)) {
	CV *cv;
	if (warnlvl < 3) {
	    cv = perl_get_cv("Carp::carp", FALSE);
	    if (cv == NULL)
		warn(SvPV(sv, len));
	}
	else {
	    cv = perl_get_cv("Carp::croak", FALSE);
	    if (cv == NULL)
		croak(SvPV(sv, len));
	}
	if (cv != NULL) {
	    PUSHMARK(sp) ;
	    XPUSHs(sv);
	    PUTBACK;
	    perl_call_sv((SV*)cv, G_DISCARD);
	}
    }

}   /* ReportOleError */

inline BOOL
CheckOleError(CPERLarg_ HV *stash, HRESULT hr, EXCEPINFO *pExcep=NULL,
	      SV *svAdd=NULL)
{
    if (FAILED(hr)) {
	ReportOleError(PERL_OBJECT_THIS_ stash, hr, pExcep, svAdd);
	return TRUE;
    }
    return FALSE;
}

SV *
CheckDestroyFunction(CPERLarg_ SV *sv, char *szMethod)
{
    /* undef */
    if (!SvOK(sv))
	return NULL;

    /* method name or CODE ref */
    if (SvPOK(sv) || (SvROK(sv) && SvTYPE(SvRV(sv)) == SVt_PVCV))
	return sv;

    warn("%s(): DESTROY must be a method name or a CODE reference", szMethod);
    DEBUGBREAK;
    return NULL;
}

void
AddToObjectChain(CPERLarg_ OBJECTHEADER *pHeader, long lMagic)
{
    dPERINTERP;

    EnterCriticalSection(&g_CriticalSection);
    pHeader->lMagic = lMagic;
    pHeader->pPrevious = NULL;
    pHeader->pNext = g_pObj;

#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    pHeader->pInterp = INTERP;
#endif

    if (g_pObj)
	g_pObj->pPrevious = pHeader;
    g_pObj = pHeader;
    LeaveCriticalSection(&g_CriticalSection);
}

void
RemoveFromObjectChain(CPERLarg_ OBJECTHEADER *pHeader)
{
    if (pHeader == NULL)
	return;

#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    PERINTERP *pInterp = pHeader->pInterp;
#endif

    EnterCriticalSection(&g_CriticalSection);
    if (pHeader->pPrevious == NULL) {
	g_pObj = pHeader->pNext;
	if (g_pObj != NULL)
	    g_pObj->pPrevious = NULL;
    }
    else if (pHeader->pNext == NULL)
	pHeader->pPrevious->pNext = NULL;
    else {
	pHeader->pPrevious->pNext = pHeader->pNext;
	pHeader->pNext->pPrevious = pHeader->pPrevious;
    }
    LeaveCriticalSection(&g_CriticalSection);
}

SV *
CreatePerlObject(CPERLarg_ HV *stash, IDispatch *pDispatch, SV *destroy)
{
    /* returns a mortal reference to a new Perl OLE object */

    if (pDispatch == NULL) {
	warn(MY_VERSION ": CreatePerlObject() No IDispatch interface");
	DEBUGBREAK;
	return &PL_sv_undef;
    }

    WINOLEOBJECT *pObj;
    HV *hvinner = newHV();
    SV *inner;
    SV *sv;
    GV **gv = (GV **) hv_fetch(stash, TIE_NAME, TIE_LEN, FALSE);
    char *szTie = szWINOLETIE;
    STRLEN len;

    if (gv != NULL && (sv = GvSV(*gv)) != NULL && SvPOK(sv))
	szTie = SvPV(sv, len);

    New(0, pObj, 1, WINOLEOBJECT);
    pObj->pDispatch = pDispatch;
    pObj->pTypeInfo = NULL;
    pObj->pEnum = NULL;
    pObj->hashTable = newHV();
    pObj->self = newHV();

    pObj->destroy = NULL;
    if (destroy != NULL) {
	if (SvPOK(destroy))
	    pObj->destroy = newSVsv(destroy);
	else if (SvROK(destroy) && SvTYPE(SvRV(destroy)) == SVt_PVCV)
	    pObj->destroy = newRV_inc(SvRV(destroy));
    }

    AddToObjectChain(PERL_OBJECT_THIS_ &pObj->header, WINOLE_MAGIC);

    DBG(("CreatePerlObject = |%lx| Class = %s Tie = %s\n", pObj,
	 HvNAME(stash), szTie));

    hv_store(hvinner, PERL_OLE_ID, PERL_OLE_IDLEN, newSViv((long)pObj), 0);
    inner = sv_bless(newRV_noinc((SV*)hvinner), gv_stashpv(szTie, TRUE));
    sv_magic((SV*)pObj->self, inner, 'P', Nullch, 0);
    SvREFCNT_dec(inner);

    return sv_2mortal(sv_bless(newRV_noinc((SV*)pObj->self), stash));

}   /* CreatePerlObject */

void
ReleasePerlObject(CPERLarg_ WINOLEOBJECT *pObj)
{
    dSP;

    if (pObj == NULL)
	return;

    /* ReleasePerlObject may be called multiple times for a single object:
     * first by Uninitialize() and then by Win32::OLE::DESTROY.
     * Make sure nothing is cleaned up twice!
     */

    if (pObj->destroy != NULL) {
	SV *self = sv_2mortal(newRV_inc((SV*)pObj->self));

	DBG(("Calling destroy method for object |%lx|\n", pObj));
	if (SvPOK(pObj->destroy)) {
	    /* Dispatch($self,$destroy,$retval); */
	    EXTEND(sp, 3);
	    PUSHMARK(sp);
	    PUSHs(self);
	    PUSHs(pObj->destroy);
	    PUSHs(sv_newmortal());
	    PUTBACK;
	    perl_call_method("Dispatch", G_DISCARD);
	}
	else {
	    PUSHMARK(sp);
	    XPUSHs(self) ;
	    PUTBACK;
	    perl_call_sv(pObj->destroy, G_DISCARD);
	}
    }

    DBG(("ReleasePerlObject |%lx|", pObj));

    if (pObj->pDispatch != NULL) {
	DBG((" pDispatch"));
	pObj->pDispatch->Release();
	pObj->pDispatch = NULL;
    }

    if (pObj->pTypeInfo != NULL) {
	DBG((" pTypeInfo"));
	pObj->pTypeInfo->Release();
	pObj->pTypeInfo = NULL;
    }

    if (pObj->pEnum != NULL) {
	DBG((" pEnum"));
	pObj->pEnum->Release();
	pObj->pEnum = NULL;
    }

    if (pObj->destroy != NULL) {
	DBG((" destroy(%d)", SvREFCNT(pObj->destroy)));
	SvREFCNT_dec(pObj->destroy);
	pObj->destroy = NULL;
    }

    if (pObj->hashTable != NULL) {
	DBG((" hashTable(%d)", SvREFCNT(pObj->hashTable)));
	SvREFCNT_dec(pObj->hashTable);
	pObj->hashTable = NULL;
    }

    DBG(("\n"));

    CoFreeUnusedLibraries();

}   /* ReleasePerlObject */

WINOLEOBJECT *
GetOleObject(CPERLarg_ SV *sv, BOOL bDESTROY=FALSE)
{
    if (sv_isobject(sv) && SvTYPE(SvRV(sv)) == SVt_PVHV) {
	SV **psv = hv_fetch((HV*)SvRV(sv), PERL_OLE_ID, PERL_OLE_IDLEN, 0);

#if (PATCHLEVEL > 4) || ((PATCHLEVEL == 4) && (SUBVERSION > 4))
	if (SvGMAGICAL(*psv))
	    mg_get(*psv);

	if (psv != NULL && SvIOK(*psv)) {
#else
	if (psv != NULL) {
#endif
	    WINOLEOBJECT *pObj = (WINOLEOBJECT*)SvIV(*psv);

	    DBG(("GetOleObject = |%lx|\n", pObj));
	    if (pObj != NULL && pObj->header.lMagic == WINOLE_MAGIC)
		if (pObj->pDispatch != NULL || bDESTROY)
		    return pObj;
	}
    }
    warn(MY_VERSION ": GetOleObject() Not a %s object", szWINOLE);
    DEBUGBREAK;
    return (WINOLEOBJECT*)NULL;
}

WINOLEENUMOBJECT *
GetOleEnumObject(CPERLarg_ SV *sv, BOOL bDESTROY=FALSE)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLEENUM)) {
	WINOLEENUMOBJECT *pEnumObj = (WINOLEENUMOBJECT*)SvIV(SvRV(sv));

	if (pEnumObj != NULL && pEnumObj->header.lMagic == WINOLEENUM_MAGIC)
	    if (pEnumObj->pEnum != NULL || bDESTROY)
		return pEnumObj;
    }
    warn(MY_VERSION ": GetOleEnumObject() Not a %s object", szWINOLEENUM);
    DEBUGBREAK;
    return (WINOLEENUMOBJECT*)NULL;
}

WINOLEVARIANTOBJECT *
GetOleVariantObject(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLEVARIANT)) {
	WINOLEVARIANTOBJECT *pVarObj = (WINOLEVARIANTOBJECT*)SvIV(SvRV(sv));

	if (pVarObj != NULL && pVarObj->header.lMagic == WINOLEVARIANT_MAGIC)
	    return pVarObj;
    }
    warn(MY_VERSION ": GetOleVariantObject() Not a %s object", szWINOLEVARIANT);
    DEBUGBREAK;
    return (WINOLEVARIANTOBJECT*)NULL;
}

SV *
CreateTypeLibObject(CPERLarg_ ITypeLib *pTypeLib, TLIBATTR *pTLibAttr)
{
    WINOLETYPELIBOBJECT *pObj;
    New(0, pObj, 1, WINOLETYPELIBOBJECT);

    pObj->pTypeLib = pTypeLib;
    pObj->pTLibAttr = pTLibAttr;

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pObj,
		     WINOLETYPELIB_MAGIC);

    return sv_bless(newRV_noinc(newSViv((IV)pObj)),
		    gv_stashpv(szWINOLETYPELIB, TRUE));
}

WINOLETYPELIBOBJECT *
GetOleTypeLibObject(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLETYPELIB)) {
	WINOLETYPELIBOBJECT *pObj = (WINOLETYPELIBOBJECT*)SvIV(SvRV(sv));

	if (pObj != NULL && pObj->header.lMagic == WINOLETYPELIB_MAGIC)
	    return pObj;
    }
    warn(MY_VERSION ": GetOleTypeLibObject() Not a %s object", szWINOLETYPELIB);
    DEBUGBREAK;
    return (WINOLETYPELIBOBJECT*)NULL;
}

SV *
CreateTypeInfoObject(CPERLarg_ ITypeInfo *pTypeInfo, TYPEATTR *pTypeAttr)
{
    WINOLETYPEINFOOBJECT *pObj;
    New(0, pObj, 1, WINOLETYPEINFOOBJECT);

    pObj->pTypeInfo = pTypeInfo;
    pObj->pTypeAttr = pTypeAttr;

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pObj,
		     WINOLETYPEINFO_MAGIC);

    return sv_bless(newRV_noinc(newSViv((IV)pObj)),
		    gv_stashpv(szWINOLETYPEINFO, TRUE));
}

WINOLETYPEINFOOBJECT *
GetOleTypeInfoObject(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLETYPEINFO)) {
	WINOLETYPEINFOOBJECT *pObj = (WINOLETYPEINFOOBJECT*)SvIV(SvRV(sv));

	if (pObj != NULL && pObj->header.lMagic == WINOLETYPEINFO_MAGIC)
	    return pObj;
    }
    warn(MY_VERSION ": GetOleTypeInfoObject() Not a %s object",
	 szWINOLETYPEINFO);
    DEBUGBREAK;
    return (WINOLETYPEINFOOBJECT*)NULL;
}

BSTR
AllocOleString(CPERLarg_ char* pStr, int length, UINT cp)
{
    int count = MultiByteToWideChar(cp, 0, pStr, length, NULL, 0);
    BSTR bstr = SysAllocStringLen(NULL, count);
    MultiByteToWideChar(cp, 0, pStr, length, bstr, count);
    return bstr;
}

HRESULT
GetHashedDispID(CPERLarg_ WINOLEOBJECT *pObj, char *buffer, STRLEN len,
		DISPID &dispID, LCID lcid, UINT cp)
{
    HRESULT hr;

    if (len == 0 || *buffer == '\0') {
	dispID = DISPID_VALUE;
	return S_OK;
    }

    SV **psv = hv_fetch(pObj->hashTable, buffer, len, 0);
    if (psv != NULL) {
	dispID = (DISPID)SvIV(*psv);
	return S_OK;
    }

    /* not there so get info and add it */
    DISPID id;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;

    pBuffer = GetWideChar(PERL_OBJECT_THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    hr = pObj->pDispatch->GetIDsOfNames(IID_NULL, &pBuffer, 1, lcid, &id);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    /* Don't call CheckOleError! Caller might retry the "unnamed" method */
    if (SUCCEEDED(hr)) {
	hv_store(pObj->hashTable, buffer, len, newSViv(id), 0);
	dispID = id;
    }
    return hr;

}   /* GetHashedDispID */

void
FetchTypeInfo(CPERLarg_ WINOLEOBJECT *pObj)
{
    unsigned int count;
    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;
    HV *stash = SvSTASH(pObj->self);

    if (pObj->pTypeInfo != NULL)
	return;

    HRESULT hr = pObj->pDispatch->GetTypeInfoCount(&count);
    if (hr == E_NOTIMPL || count == 0) {
	DBG(("GetTypeInfoCount returned %u (count=%d)", hr, count));
	return;
    }

    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
	warn(MY_VERSION ": FetchTypeInfo() GetTypeInfoCount failed");
	DEBUGBREAK;
	return;
    }

    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	return;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	return;
    }

    if (pTypeAttr->typekind != TKIND_DISPATCH) {
	int cImplTypes = pTypeAttr->cImplTypes;
	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeAttr = NULL;

	for (int i=0 ; i < cImplTypes ; ++i) {
	    HREFTYPE hreftype;
	    ITypeInfo *pRefTypeInfo;

	    hr = pTypeInfo->GetRefTypeOfImplType(i, &hreftype);
	    if (FAILED(hr))
		break;

	    hr = pTypeInfo->GetRefTypeInfo(hreftype, &pRefTypeInfo);
	    if (FAILED(hr))
		break;

	    hr = pRefTypeInfo->GetTypeAttr(&pTypeAttr);
	    if (FAILED(hr)) {
		pRefTypeInfo->Release();
		break;
	    }

	    if (pTypeAttr->typekind == TKIND_DISPATCH) {
		pTypeInfo->Release();
		pTypeInfo = pRefTypeInfo;
		break;
	    }

	    pRefTypeInfo->ReleaseTypeAttr(pTypeAttr);
	    pRefTypeInfo->Release();
	    pTypeAttr = NULL;
	}
    }

    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	return;
    }

    if (pTypeAttr != NULL) {
	if (pTypeAttr->typekind == TKIND_DISPATCH) {
	    pObj->cFuncs = pTypeAttr->cFuncs;
	    pObj->cVars = pTypeAttr->cVars;
	    pObj->PropIndex = 0;
	    pObj->pTypeInfo = pTypeInfo;
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	if (pObj->pTypeInfo == NULL)
	    pTypeInfo->Release();
    }

}   /* FetchTypeInfo */

SV *
NextPropertyName(CPERLarg_ WINOLEOBJECT *pObj)
{
    HRESULT hr;
    unsigned int cName;
    BSTR bstr;

    if (pObj->pTypeInfo == NULL)
	return &PL_sv_undef;

    HV *stash = SvSTASH(pObj->self);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    while (pObj->PropIndex < pObj->cFuncs+pObj->cVars) {
	ULONG index = pObj->PropIndex++;
	/* Try all the INVOKE_PROPERTYGET functions first */
	if (index < pObj->cFuncs) {
	    FUNCDESC *pFuncDesc;

	    hr = pObj->pTypeInfo->GetFuncDesc(index, &pFuncDesc);
	    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
		continue;

	    if (!(pFuncDesc->funckind & FUNC_DISPATCH) ||
		!(pFuncDesc->invkind & INVOKE_PROPERTYGET) ||
	        (pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED |
					  FUNCFLAG_FHIDDEN |
					  FUNCFLAG_FNONBROWSABLE)))
	    {
		pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		continue;
	    }

	    hr = pObj->pTypeInfo->GetNames(pFuncDesc->memid, &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
	    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) ||
		cName == 0 || bstr == NULL)
	    {
		continue;
	    }

	    SV *sv = sv_setwide(PERL_OBJECT_THIS_ NULL, bstr, cp);
	    SysFreeString(bstr);
	    return sv;
	}
	/* Now try the VAR_DISPATCH kind variables used by older OLE versions */
	else {
	    VARDESC *pVarDesc;

	    index -= pObj->cFuncs;
	    hr = pObj->pTypeInfo->GetVarDesc(index, &pVarDesc);
	    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
		continue;

	    if (!(pVarDesc->varkind & VAR_DISPATCH) ||
		(pVarDesc->wVarFlags & (VARFLAG_FRESTRICTED |
					VARFLAG_FHIDDEN |
					VARFLAG_FNONBROWSABLE)))
	    {
		pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
		continue;
	    }

	    hr = pObj->pTypeInfo->GetNames(pVarDesc->memid, &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
	    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) ||
		cName == 0 || bstr == NULL)
	    {
		continue;
	    }

	    SV *sv = sv_setwide(PERL_OBJECT_THIS_ NULL, bstr, cp);
	    SysFreeString(bstr);
	    return sv;
	}
    }
    return &PL_sv_undef;

}   /* NextPropertyName */

HV *
GetDocumentation(CPERLarg_ BSTR bstrName, BSTR bstrDocString,
		 DWORD dwHelpContext, BSTR bstrHelpFile)
{
    HV *hv = newHV();
    char szStr[OLE_BUF_SIZ];
    char *pszStr;
    // XXX use correct codepage ???
    UINT cp = CP_ACP;

    pszStr = GetMultiByte(PERL_OBJECT_THIS_ bstrName,
			  szStr, sizeof(szStr), cp);
    hv_store(hv, "Name", 4, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(PERL_OBJECT_THIS_ pszStr, szStr);
    SysFreeString(bstrName);

    pszStr = GetMultiByte(PERL_OBJECT_THIS_ bstrDocString,
			  szStr, sizeof(szStr), cp);
    hv_store(hv, "DocString", 9, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(PERL_OBJECT_THIS_ pszStr, szStr);
    SysFreeString(bstrDocString);

    pszStr = GetMultiByte(PERL_OBJECT_THIS_ bstrHelpFile,
			  szStr, sizeof(szStr), cp);
    hv_store(hv, "HelpFile", 8, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(PERL_OBJECT_THIS_ pszStr, szStr);
    SysFreeString(bstrHelpFile);

    hv_store(hv, "HelpContext", 11, newSViv(dwHelpContext), 0);

    return hv;

}   /* GetDocumentation */

HRESULT
TranslateTypeDesc(CPERLarg_ TYPEDESC *pTypeDesc, WINOLETYPEINFOOBJECT *pObj,
		  AV *av)
{
    HRESULT hr = S_OK;
    SV *sv = NULL;

    if (pTypeDesc->vt == VT_USERDEFINED) {
	ITypeInfo *pTypeInfo;
	TYPEATTR  *pTypeAttr;
	hr = pObj->pTypeInfo->GetRefTypeInfo(pTypeDesc->hreftype, &pTypeInfo);
	if (SUCCEEDED(hr)) {
	    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	    if (SUCCEEDED(hr)) {
		sv = CreateTypeInfoObject(PERL_OBJECT_THIS_ pTypeInfo,
					  pTypeAttr);
	    }
	    else
		pTypeInfo->Release();
	}
	if (!sv)
	    sv = newSVsv(&PL_sv_undef);

    }
    else if (pTypeDesc->vt == VT_CARRAY) {
	// XXX to be done
	sv = newSViv(pTypeDesc->vt);
    }
    else
	sv = newSViv(pTypeDesc->vt);

    av_push(av, sv);

    if (pTypeDesc->vt == VT_PTR || pTypeDesc->vt == VT_SAFEARRAY)
	hr = TranslateTypeDesc(PERL_OBJECT_THIS_ pTypeDesc->lptdesc, pObj, av);

    return hr;
}

HV *
TranslateElemDesc(CPERLarg_ ELEMDESC *pElemDesc, WINOLETYPEINFOOBJECT *pObj,
		  HV *olestash)
{
    HV *hv = newHV();

    AV *av = newAV();
    TranslateTypeDesc(PERL_OBJECT_THIS_  &pElemDesc->tdesc, pObj, av);
    hv_store(hv, "vt", 2, newRV_noinc((SV*)av), 0);

    USHORT wParamFlags = pElemDesc->paramdesc.wParamFlags;
    hv_store(hv, "wParamFlags", 11, newSViv(wParamFlags), 0);

    USHORT wMask = PARAMFLAG_FOPT|PARAMFLAG_FHASDEFAULT;
    if ((wParamFlags & wMask) == wMask) {
	PARAMDESCEX *pParamDescEx = pElemDesc->paramdesc.pparamdescex;
	hv_store(hv, "cBytes", 6, newSViv(pParamDescEx->cBytes), 0);
	// XXX should be stored as a Win32::OLE::Variant object ?
	SV *sv = newSVpv("",0);
	SetSVFromVariantEx(PERL_OBJECT_THIS_ &pParamDescEx->varDefaultValue,
			   sv, olestash);
	hv_store(hv, "varDefaultValue", 15, sv, 0);
    }

    return hv;

}   /* TranslateElemDesc */

IEnumVARIANT *
CreateEnumVARIANT(CPERLarg_ WINOLEOBJECT *pObj)
{
    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    HRESULT hr;
    IEnumVARIANT *pEnum = NULL;

    VariantInit(&result);
    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = 0;

    HV *stash = SvSTASH(pObj->self);
    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);

    Zero(&excepinfo, 1, EXCEPINFO);
    hr = pObj->pDispatch->Invoke(DISPID_NEWENUM, IID_NULL,
			    lcid, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			    &dispParams, &result, &excepinfo, &argErr);
    if (SUCCEEDED(hr)) {
	if (V_VT(&result) == VT_UNKNOWN)
	    hr = V_UNKNOWN(&result)->QueryInterface(IID_IEnumVARIANT,
						    (void**)&pEnum);
	else if (V_VT(&result) == VT_DISPATCH)
	    hr = V_DISPATCH(&result)->QueryInterface(IID_IEnumVARIANT,
						     (void**)&pEnum);
    }
    VariantClear(&result);
    CheckOleError(PERL_OBJECT_THIS_ stash, hr, &excepinfo);
    return pEnum;

}   /* CreateEnumVARIANT */

SV *
NextEnumElement(CPERLarg_ IEnumVARIANT *pEnum, HV *stash)
{
    HRESULT hr = S_OK;
    SV *sv = &PL_sv_undef;
    VARIANT variant;

    VariantInit(&variant);
    if (SUCCEEDED(pEnum->Next(1, &variant, NULL))) {
	sv = newSVpv("",0);
	hr = SetSVFromVariantEx(PERL_OBJECT_THIS_ &variant, sv, stash);
    }
    VariantClear(&variant);
    if (FAILED(hr)) {
        SvREFCNT_dec(sv);
	sv = &PL_sv_undef;
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
    }
    return sv;

}   /* NextEnumElement */

SV *
SetSVFromGUID(CPERLarg_ REFGUID rguid)
{
    dSP;
    SV *sv = newSVsv(&PL_sv_undef);
    CV *cv = perl_get_cv("Win32::COM::GUID::new", FALSE);

    if (cv == NULL) {
	OLECHAR wszGUID[80];
	int len = StringFromGUID2(rguid, wszGUID, sizeof(wszGUID)/sizeof(OLECHAR));
	if (len > 0) {
	    wszGUID[len-2] = (OLECHAR) 0;
	    sv_setwide(PERL_OBJECT_THIS_ sv, wszGUID+1, CP_ACP);
	}
    }
    else
    {
	EXTEND(sp, 2);
	PUSHMARK(sp);
	PUSHs(sv_2mortal(newSVpv("Win32::COM::GUID", 0)));
	PUSHs(sv_2mortal(newSVpv((char*)&rguid, sizeof(GUID))));
	PUTBACK;
	int count = perl_call_sv((SV*)cv, G_SCALAR);
	SPAGAIN;
	if (count == 1)
	    sv_setsv(sv, POPs);
	PUTBACK;
    }
    return sv;
}

HRESULT
SetVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp)
{
    HRESULT hr = S_OK;
    VariantClear(pVariant);

    /* XXX requirement to call mg_get() may change in Perl > 5.005 */
    if (SvGMAGICAL(sv))
	mg_get(sv);

    /* Objects */
    if (SvROK(sv)) {
	if (sv_derived_from(sv, szWINOLE)) {
	    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ sv);
	    if (pObj == NULL)
		hr = E_POINTER;
	    else {
		pObj->pDispatch->AddRef();
		V_VT(pVariant) = VT_DISPATCH;
		V_DISPATCH(pVariant) = pObj->pDispatch;
	    }
	    return hr;
	}

	if (sv_derived_from(sv, szWINOLEVARIANT)) {
	    WINOLEVARIANTOBJECT *pVarObj =
		GetOleVariantObject(PERL_OBJECT_THIS_ sv);

	    if (pVarObj == NULL)
		hr = E_POINTER;
	    else {
		/* XXX Should we use VariantCopyInd? */
		hr = VariantCopy(pVariant, &pVarObj->variant);
	    }
	    return hr;
	}

	sv = SvRV(sv);
    }

    /* Arrays */
    if (SvTYPE(sv) == SVt_PVAV) {
	IV index;
	IV dim = 1;
	IV maxdim = 2;
	AV **pav;
	long *pix;
	long *plen;
	SAFEARRAYBOUND *psab;

	New(0, pav, maxdim, AV*);
	New(0, pix, maxdim, long);
	New(0, plen, maxdim, long);
	New(0, psab, maxdim, SAFEARRAYBOUND);

	pav[0] = (AV*)sv;
	pix[0] = 0;
	plen[0] = av_len(pav[0])+1;
	psab[0].cElements = plen[0];
	psab[0].lLbound = 0;

	/* Depth first walk through to determine number of dimensions */
	for (index = 0 ; index >= 0 ; ) {
	    SV **psv = av_fetch(pav[index], pix[index], FALSE);

	    if (psv != NULL && SvROK(*psv) && SvTYPE(SvRV(*psv)) == SVt_PVAV) {
		if (++index >= maxdim) {
		    maxdim *= 2;
		    Renew(pav, maxdim, AV*);
		    Renew(pix, maxdim, long);
		    Renew(plen, maxdim, long);
		    Renew(psab, maxdim, SAFEARRAYBOUND);
		}

		pav[index] = (AV*)SvRV(*psv);
		pix[index] = 0;
		plen[index] = av_len(pav[index])+1;

		if (index < dim) {
		    if (plen[index] > psab[index].cElements)
			psab[index].cElements = plen[index];
		}
		else {
		    dim = index+1;
		    psab[index].cElements = plen[index];
		    psab[index].lLbound = 0;
		}
		continue;
	    }

	    while (index >= 0) {
		if (++pix[index] < plen[index])
		    break;
		--index;
	    }
	}

	/* Create and fill VARIANT array */
	SAFEARRAY *psa = SafeArrayCreate(VT_VARIANT, dim, psab);
	if (psa == NULL)
	    hr = E_OUTOFMEMORY;
	else
	    hr = SafeArrayLock(psa);

	if (SUCCEEDED(hr)) {
	    pav[0] = (AV*)sv;
	    plen[0] = av_len(pav[0])+1;
	    Zero(pix, dim, long);

	    for (index = 0 ; index >= 0 ; ) {
		SV **psv = av_fetch(pav[index], pix[index], FALSE);

		if (psv != NULL) {
		    if (SvROK(*psv) && SvTYPE(SvRV(*psv)) == SVt_PVAV) {
			++index;
			pav[index] = (AV*)SvRV(*psv);
			pix[index] = 0;
			plen[index] = av_len(pav[index])+1;
			continue;
		    }

		    if (SvOK(*psv)) {
			VARIANT *pElement;
			hr = SafeArrayPtrOfIndex(psa, pix, (void**)&pElement);
			if (SUCCEEDED(hr))
			    hr = SetVariantFromSV(PERL_OBJECT_THIS_ *psv,
						  pElement, cp);
			if (FAILED(hr))
			    break;
		    }
		}

		while (index >= 0) {
		    if (++pix[index] < plen[index])
			break;
		    pix[index--] = 0;
		}
	    }
	    hr = SafeArrayUnlock(psa);
	}

	Safefree(pav);
	Safefree(pix);
	Safefree(plen);
	Safefree(psab);

	if (SUCCEEDED(hr)) {
	    V_VT(pVariant) = VT_VARIANT | VT_ARRAY;
	    V_ARRAY(pVariant) = psa;
	}
	else if (psa != NULL)
	    SafeArrayDestroy(psa);

	return hr;
    }

    /* Scalars */
    if (SvIOK(sv)) {
	V_VT(pVariant) = VT_I4;
	V_I4(pVariant) = SvIV(sv);
    }
    else if (SvNOK(sv)) {
	V_VT(pVariant) = VT_R8;
	V_R8(pVariant) = SvNV(sv);
    }
    else if (SvPOK(sv)) {
	STRLEN len;
	char *ptr = SvPV(sv, len);
	V_VT(pVariant) = VT_BSTR;
	V_BSTR(pVariant) = AllocOleString(PERL_OBJECT_THIS_ ptr, len, cp);
    }
    else {
	V_VT(pVariant) = VT_ERROR;
	V_ERROR(pVariant) = DISP_E_PARAMNOTFOUND;
    }

    return hr;

}   /* SetVariantFromSV */

HRESULT
AssignVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant, HV *stash)
{
    /* This function is similar to SetVariantFromSV except that
     * it does NOT choose the variant type itself.
     */
    HRESULT hr = S_OK;
    VARTYPE vt = V_VT(pVariant);

#   define ASSIGN(vartype,perltype)                           \
        if (vt & VT_BYREF) {                                  \
            *V_##vartype##REF(pVariant) = Sv##perltype##(sv); \
        } else {                                              \
            V_##vartype(pVariant) = Sv##perltype##(sv);       \
        }

    /* XXX requirement to call mg_get() may change in Perl > 5.005 */
    if (sv != NULL && SvGMAGICAL(sv))
	mg_get(sv);

    if (vt & VT_ARRAY) {
	SAFEARRAY *psa;
	if (V_ISBYREF(pVariant))
	    psa = *V_ARRAYREF(pVariant);
	else
	    psa = V_ARRAY(pVariant);

	UINT cDims = SafeArrayGetDim(psa);
	VARTYPE vt_base = vt & ~VT_BYREF & ~VT_ARRAY;
	if (vt_base != VT_UI1 || cDims != 1 || !SvPOK(sv)) {
	    warn(MY_VERSION ": AssignVariantFromSV() cannot assign to "
		 "VT_ARRAY variant");
	    return E_INVALIDARG;
	}

	char *pDest;
	STRLEN len;
	char *pSrc = SvPV(sv, len);
	HRESULT hr = SafeArrayAccessData(psa, (void**)&pDest);
	if (FAILED(hr))
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	else {
	    long lLower, lUpper;
	    SafeArrayGetLBound(psa, 1, &lLower);
	    SafeArrayGetUBound(psa, 1, &lUpper);

	    long lLength = 1 + lUpper-lLower;
	    len = min(len, lLength);
	    memcpy(pDest, pSrc, len);
	    if (lLength > len)
		memset(pDest+len, 0, lLength-len);

	    SafeArrayUnaccessData(psa);
	}
	return hr;
    }

    switch(vt & ~VT_BYREF) {
    case VT_EMPTY:
    case VT_NULL:
	break;

    case VT_I2:
	ASSIGN(I2, IV);
	break;

    case VT_I4:
	ASSIGN(I4, IV);
	break;

    case VT_R4:
	ASSIGN(R4, NV);
	break;

    case VT_R8:
	ASSIGN(R8, NV);
	break;

    case VT_CY:
    case VT_DATE:
    {
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
				lcidDefault);
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);
	STRLEN len;
	char *ptr = SvPV(sv, len);

	V_VT(pVariant) = VT_BSTR;
	V_BSTR(pVariant) = AllocOleString(PERL_OBJECT_THIS_ ptr, len, cp);
	/* XXX VT_BYREF ??? */
	VariantChangeTypeEx(pVariant, pVariant, lcid,0, vt);
	break;
    }

    case VT_BSTR:
    {
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);
	STRLEN len;
	char *ptr = SvPV(sv, len);
	BSTR bstr = AllocOleString(PERL_OBJECT_THIS_ ptr, len, cp);

	if (vt & VT_BYREF) {
	    SysFreeString(*V_BSTRREF(pVariant));
	    *V_BSTRREF(pVariant) = bstr;
	}
	else {
	    SysFreeString(V_BSTR(pVariant));
	    V_BSTR(pVariant) = bstr;
	}
	break;
    }

    case VT_DISPATCH:
    {
	/* Argument MUST be a valid Perl OLE object! */
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ sv);
	if (pObj != NULL) {
	    pObj->pDispatch->AddRef();
	    if (vt & VT_BYREF) {
		if (*V_DISPATCHREF(pVariant) != NULL)
		    (*V_DISPATCHREF(pVariant))->Release();
		*V_DISPATCHREF(pVariant) = pObj->pDispatch;
	    }
	    else {
		if (V_DISPATCH(pVariant) != NULL)
		    V_DISPATCH(pVariant)->Release();
		V_DISPATCH(pVariant) = pObj->pDispatch;
	    }
	}
	break;
    }

    case VT_ERROR:
	ASSIGN(ERROR, IV);
	break;

    case VT_BOOL:
	if (vt & VT_BYREF)
	    *V_BOOLREF(pVariant) = SvTRUE(sv) ? VARIANT_TRUE : VARIANT_FALSE;
	else
	    V_BOOL(pVariant) = SvTRUE(sv) ? VARIANT_TRUE : VARIANT_FALSE;
	break;

    case VT_VARIANT:
	if (vt & VT_BYREF) {
	    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
				  cpDefault);
	    hr = SetVariantFromSV(PERL_OBJECT_THIS_ sv,
				  V_VARIANTREF(pVariant), cp);
	}
	else {
	    warn(MY_VERSION ": AssignVariantFromSV() with invalid type: "
		 "VT_VARIANT without VT_BYREF");
	    hr = E_INVALIDARG;
	}
	break;

    case VT_UNKNOWN:
    {
	/* Argument MUST be a valid Perl OLE object! */
	/* Query IUnknown interface to allow identity tests */
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ sv);
	if (pObj != NULL) {
	    IUnknown *punk;
	    hr = pObj->pDispatch->QueryInterface(IID_IUnknown, (void**)&punk);
	    if (!CheckOleError(PERL_OBJECT_THIS_ SvSTASH(pObj->self), hr)) {
		if (vt & VT_BYREF) {
		    if (*V_UNKNOWNREF(pVariant) != NULL)
			(*V_UNKNOWNREF(pVariant))->Release();
		    *V_UNKNOWNREF(pVariant) = punk;
		}
		else {
		    if (V_UNKNOWN(pVariant) != NULL)
			V_UNKNOWN(pVariant)->Release();
		    V_UNKNOWN(pVariant) = punk;
		}
	    }
	}
	break;
    }

    case VT_UI1:
	if (SvIOK(sv)) {
	    ASSIGN(UI1, IV);
	}
	else {
	    STRLEN len;
	    char *ptr = SvPV(sv, len);
	    if (vt & VT_BYREF)
		*V_UI1REF(pVariant) = *ptr;
	    else
		V_UI1(pVariant) = *ptr;
	}
	break;

    default:
	warn(MY_VERSION " AssignVariantFromSV() cannot assign to "
	     "vt=0x%x", vt);
	hr = E_INVALIDARG;
    }

    return hr;
#   undef ASSIGN
}   /* AssignVariantFromSV */

HRESULT
SetSVFromVariantEx(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash,
		   BOOL bByRefObj)
{
    HRESULT hr = S_OK;
    VARTYPE vt = V_VT(pVariant);

#   define SET(perltype,vartype)                                 \
        if (vt & VT_BYREF) {                                     \
            sv_set##perltype##(sv, *V_##vartype##REF(pVariant)); \
        } else {                                                 \
            sv_set##perltype##(sv, V_##vartype##(pVariant));     \
        }

    sv_setsv(sv, &PL_sv_undef);

    if (V_ISBYREF(pVariant) && bByRefObj) {
	WINOLEVARIANTOBJECT *pVarObj;
	Newz(0, pVarObj, 1, WINOLEVARIANTOBJECT);
	VariantInit(&pVarObj->variant);
	VariantInit(&pVarObj->byref);
	hr = VariantCopy(&pVarObj->variant, pVariant);
	if (FAILED(hr)) {
	    Safefree(pVarObj);
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	}

	AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pVarObj,
			 WINOLEVARIANT_MAGIC);
	STRLEN len;
	SV *classname = newSVpv(HvNAME(stash), 0);
	sv_catpvn(classname, "::Variant", 9);
	sv_setref_pv(sv, SvPV(classname, len), pVarObj);
	SvREFCNT_dec(classname);
	return hr;
    }

    if (V_ISARRAY(pVariant)) {
	SAFEARRAY *psa = V_ISBYREF(pVariant) ? *V_ARRAYREF(pVariant)
	                                     : V_ARRAY(pVariant);
	AV **pav;
	IV index;
	long *pArrayIndex, *pLowerBound, *pUpperBound;
	VARIANT variant;

	int dim = SafeArrayGetDim(psa);

	VariantInit(&variant);
	V_VT(&variant) = (vt & ~VT_ARRAY) | VT_BYREF;

	/* convert 1-dim UI1 ARRAY to simple SvPV */
	if (dim == 1 && (vt & ~VT_ARRAY & ~VT_BYREF) == VT_UI1) {
	    char *pStr;
	    long lLower, lUpper;

	    SafeArrayGetLBound(psa, 1, &lLower);
	    SafeArrayGetUBound(psa, 1, &lUpper);
	    hr = SafeArrayAccessData(psa, (void**)&pStr);
	    if (SUCCEEDED(hr)) {
		sv_setpvn(sv, pStr, lUpper-lLower+1);
		SafeArrayUnaccessData(psa);
	    }

	    return hr;
	}

	New(0, pArrayIndex, dim, long);
	New(0, pLowerBound, dim, long);
	New(0, pUpperBound, dim, long);
	New(0, pav,         dim, AV *);

	for(index = 0; index < dim; ++index) {
	    pav[index] = newAV();
	    SafeArrayGetLBound(psa, index+1, &pLowerBound[index]);
	    SafeArrayGetUBound(psa, index+1, &pUpperBound[index]);
	}

	Copy(pLowerBound, pArrayIndex, dim, long);

	hr = SafeArrayLock(psa);
	if (SUCCEEDED(hr)) {
	    while (index >= 0) {
		hr = SafeArrayPtrOfIndex(psa, pArrayIndex, &V_BYREF(&variant));
		if (FAILED(hr))
		    break;

		SV *val = newSVpv("",0);
		hr = SetSVFromVariantEx(PERL_OBJECT_THIS_ &variant, val, stash);
		if (FAILED(hr)) {
		    SvREFCNT_dec(val);
		    break;
		}
		av_push(pav[dim-1], val);

		for (index = dim-1 ; index >= 0 ; --index) {
		    if (++pArrayIndex[index] <= pUpperBound[index])
			break;

		    pArrayIndex[index] = pLowerBound[index];
		    if (index > 0) {
			av_push(pav[index-1], newRV_noinc((SV*)pav[index]));
			pav[index] = newAV();
		    }
		}
	    }

	    /* preserve previous error code */
	    HRESULT hr2 = SafeArrayUnlock(psa);
	    if (SUCCEEDED(hr))
		hr = hr2;
	}

	for (index = 1 ; index < dim ; ++index)
	    SvREFCNT_dec((SV*)pav[index]);

	if (SUCCEEDED(hr))
	    sv_setsv(sv, sv_2mortal(newRV_noinc((SV*)*pav)));
	else
	    SvREFCNT_dec((SV*)*pav);

	Safefree(pArrayIndex);
	Safefree(pLowerBound);
	Safefree(pUpperBound);
	Safefree(pav);

	return hr;
    }

    while (vt == (VT_VARIANT|VT_BYREF)) {
	pVariant = V_VARIANTREF(pVariant);
	vt = V_VT(pVariant);
    }

    switch(vt & ~VT_BYREF) {
    case VT_VARIANT: /* invalid, should never happen */
    case VT_EMPTY:
    case VT_NULL:
	/* return "undef" */
	break;

    case VT_UI1:
	SET(iv, UI1);
	break;

    case VT_I2:
	SET(iv, I2);
	break;

    case VT_I4:
	SET(iv, I4);
	break;

    case VT_R4:
	SET(nv, R4);
	break;

    case VT_R8:
	SET(nv, R8);
	break;

    case VT_BSTR:
    {
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);

	if (V_ISBYREF(pVariant))
	    sv_setwide(PERL_OBJECT_THIS_ sv, *V_BSTRREF(pVariant), cp);
	else
	    sv_setwide(PERL_OBJECT_THIS_ sv, V_BSTR(pVariant), cp);

	break;
    }

    case VT_ERROR:
	SET(iv, ERROR);
	break;

    case VT_BOOL:
	if (V_ISBYREF(pVariant))
	    sv_setiv(sv, *V_BOOLREF(pVariant) ? 1 : 0);
	else
	    sv_setiv(sv, V_BOOL(pVariant) ? 1 : 0);
	break;

    case VT_DISPATCH:
    {
	IDispatch *pDispatch;

	if (V_ISBYREF(pVariant))
	    pDispatch = *V_DISPATCHREF(pVariant);
	else
	    pDispatch = V_DISPATCH(pVariant);

	if (pDispatch != NULL ) {
	    pDispatch->AddRef();
	    sv_setsv(sv, CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch,
					  NULL));
	}
	break;
    }

    case VT_UNKNOWN:
    {
	IUnknown *punk;
	IDispatch *pDispatch;

	if (V_ISBYREF(pVariant))
	    punk = *V_UNKNOWNREF(pVariant);
	else
	    punk = V_UNKNOWN(pVariant);

	if (punk != NULL &&
	    SUCCEEDED(punk->QueryInterface(IID_IDispatch, (void**)&pDispatch)))
	{
	    sv_setsv(sv, CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch,
					  NULL));
	}
	break;
    }

    case VT_DATE:
    case VT_CY:
    default:
    {
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
				lcidDefault);
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);
	VARIANT variant;

	VariantInit(&variant);
	hr = VariantChangeTypeEx(&variant, pVariant, lcid, 0, VT_BSTR);
	if (SUCCEEDED(hr) && V_VT(&variant) == VT_BSTR)
	    sv_setwide(PERL_OBJECT_THIS_ sv, V_BSTR(&variant), cp);
	VariantClear(&variant);
	break;
    }
    }

    return hr;
#   undef SET
}   /* SetSVFromVariantEx */

HRESULT
SetSVFromVariant(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash)
{
    return SetSVFromVariantEx(PERL_OBJECT_THIS_ pVariant, sv, stash);
}

inline void
SpinMessageLoop(void)
{
    MSG msg;

    DBG(("SpinMessageLoop\n"));
    while(PeekMessage(&msg,NULL,NULL,NULL,PM_REMOVE)) {
	TranslateMessage(&msg);
	DispatchMessage(&msg);
    }

}   /* SpinMessageLoop */

void
Initialize(CPERLarg_ HV *stash, DWORD dwCoInit=COINIT_MULTITHREADED)
{
    dPERINTERP;

    DBG(("Initialize\n"));
    EnterCriticalSection(&g_CriticalSection);

    if (!g_bInitialized)
    {
	HRESULT hr = S_OK;

	g_pfnCoUninitialize = NULL;
	g_bInitialized = TRUE;

	DBG(("(Co|Ole)Initialize(Ex)?\n"));

	if (dwCoInit == COINIT_OLEINITIALIZE) {
	    hr = OleInitialize(NULL);
	    if (SUCCEEDED(hr))
		g_pfnCoUninitialize = &OleUninitialize;
	}
	else {
	    if (g_pfnCoInitializeEx == NULL)
		hr = CoInitialize(NULL);
	    else
		hr = g_pfnCoInitializeEx(NULL, dwCoInit);

	    if (SUCCEEDED(hr))
		g_pfnCoUninitialize = &CoUninitialize;
	}

	if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
    }

    LeaveCriticalSection(&g_CriticalSection);

}   /* Initialize */

void
Uninitialize(CPERLarg_ PERINTERP *pInterp, int magic=0)
{
    /* This function is called during Perl interpreter cleanup after all objects
     * have already been destroyed. Do NOT access Perl data structures! */

    DBG(("Uninitialize\n"));
    EnterCriticalSection(&g_CriticalSection);
    if (g_bInitialized) {
	while (g_pObj != NULL) {
	    DBG(("Zombiefy object |%lx|\n", g_pObj));

	    switch (g_pObj->lMagic)
	    {
	    case WINOLE_MAGIC:
		ReleasePerlObject(PERL_OBJECT_THIS_ (WINOLEOBJECT*)g_pObj);
		break;

	    case WINOLEENUM_MAGIC:
	    {
		WINOLEENUMOBJECT *pEnumObj = (WINOLEENUMOBJECT*)g_pObj;
		if (pEnumObj->pEnum != NULL) {
		    pEnumObj->pEnum->Release();
		    pEnumObj->pEnum = NULL;
		}
		break;
	    }

	    case WINOLEVARIANT_MAGIC:
	    {
		WINOLEVARIANTOBJECT *pVarObj = (WINOLEVARIANTOBJECT*)g_pObj;
		VariantClear(&pVarObj->byref);
		VariantClear(&pVarObj->variant);
		break;
	    }

	    case WINOLETYPELIB_MAGIC:
	    {
		WINOLETYPELIBOBJECT *pObj = (WINOLETYPELIBOBJECT*)g_pObj;
		if (pObj->pTypeLib != NULL) {
		    pObj->pTypeLib->Release();
		    pObj->pTypeLib = NULL;
		}
		break;
	    }

	    case WINOLETYPEINFO_MAGIC:
	    {
		WINOLETYPEINFOOBJECT *pObj = (WINOLETYPEINFOOBJECT*)g_pObj;
		if (pObj->pTypeInfo != NULL) {
		    pObj->pTypeInfo->Release();
		    pObj->pTypeInfo = NULL;
		}
		break;
	    }

	    default:
		DBG(("Unknown magic number: %08lx", g_pObj->lMagic));
		break;
	    }
	    g_pObj = g_pObj->pNext;
	}

	SpinMessageLoop();
	DBG(("CoUninitialize\n"));
	if (g_pfnCoUninitialize != NULL)
	    g_pfnCoUninitialize();
	g_bInitialized = FALSE;
    }
    LeaveCriticalSection(&g_CriticalSection);

    if (magic == WINOLE_MAGIC) {
    // Yes, we DO leak the critical section and memory block for earlier
    // versions of Perl (because they might still be referenced during the
    // global object destruction phase).
#if (PATCHLEVEL > 4) || (SUBVERSION >= 68)
	DeleteCriticalSection(&g_CriticalSection);
	if (g_hOLE32 != NULL)
	    FreeLibrary(g_hOLE32);
#   if defined(MULTIPLICITY) || defined(PERL_OBJECT)
	Safefree(pInterp);
#   endif
#endif
	DBG(("Interpreter exit\n"));
    }

}   /* Uninitialize */

static void
AtExit(CPERLarg_ void *pVoid)
{
    Uninitialize(PERL_OBJECT_THIS_ (PERINTERP*)pVoid, WINOLE_MAGIC);
    DBG(("AtExit done\n"));

}   /* AtExit */

void
Bootstrap(CPERLarg)
{
#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    PERINTERP *pInterp;
    New(0, pInterp, 1, PERINTERP);

#   if (PATCHLEVEL == 4) && (SUBVERSION < 68)
    SV *sv = perl_get_sv(MY_VERSION, TRUE);
#   else
    SV *sv = *hv_fetch(PL_modglobal, MY_VERSION, sizeof(MY_VERSION)-1, TRUE);
#   endif

    if (SvOK(sv))
	warn(MY_VERSION ": Per-interpreter data already set");

    sv_setiv(sv, (IV)pInterp);
#endif

    g_pObj = NULL;
    g_bInitialized = FALSE;
    InitializeCriticalSection(&g_CriticalSection);

    g_hOLE32 = LoadLibrary("OLE32");
    g_pfnCoInitializeEx = NULL;
    g_pfnCoCreateInstanceEx = NULL;
    if (g_hOLE32 != NULL) {
	g_pfnCoInitializeEx = (FNCOINITIALIZEEX*)
	    GetProcAddress(g_hOLE32, "CoInitializeEx");
	g_pfnCoCreateInstanceEx = (FNCOCREATEINSTANCEEX*)
	    GetProcAddress(g_hOLE32, "CoCreateInstanceEx");
    }

#if (PATCHLEVEL == 4) && (SUBVERSION < 68)
    SV *cmd = sv_2mortal(newSVpv("",0));
    sv_setpvf(cmd, "END { %s->Uninitialize(%d); }", szWINOLE, WINOLE_MAGIC );
    perl_eval_sv(cmd, TRUE);
#else
    perl_atexit(AtExit, INTERP);
#endif

}   /* Bootstrap */

BOOL
CallObjectMethod(CPERLarg_ SV **mark, I32 ax, I32 items, char *pszMethod)
{
    /* If the 1st arg on the stack is a Win32::OLE object then the method
     * is called as an object method through Win32::OLE::Dispatch (like
     * the AUTOLOAD does) and CallObjectMethod returns TRUE. In this case
     * the caller should return immediately. Otherwise it should check the
     * parameters on the stack and implement its class method functionality.
     */
    dSP;

    if (items == 0)
	return FALSE;

    if (!sv_isobject(ST(0)) || !sv_derived_from(ST(0), szWINOLE))
	return FALSE;

    SV *retval = sv_newmortal();

    /* Dispatch must be called as: Dispatch($self,$method,$retval,@params),
     * so move all stack entries after the object ref up to make room for
     * the method name and return value.
     */
    PUSHMARK(mark);
    EXTEND(sp,2);
    for (I32 item = 1 ; item < items ; ++item)
	ST(2+items-item) = ST(items-item);
    sp += 2;

    ST(1) = sv_2mortal(newSVpv(pszMethod,0));
    ST(2) = retval;

    PUTBACK;
    perl_call_method("Dispatch", G_DISCARD);
    SPAGAIN;

    PUSHs(retval);
    PUTBACK;

    return TRUE;

}   /* CallObjectMethod */

#if defined (__cplusplus)
}
#endif

/*##########################################################################*/

MODULE = Win32::OLE		PACKAGE = Win32::OLE

PROTOTYPES: DISABLE

BOOT:
    Bootstrap(PERL_OBJECT_THIS);

void
Initialize(...)
ALIAS:
    Uninitialize = 1
    SpinMessageLoop = 2
PPCODE:
{
    char *paszMethod[] = {"Initialize", "Uninitialize", "SpinMessageLoop"};

    if (CallObjectMethod(PERL_OBJECT_THIS_ mark, ax, items, paszMethod[ix]))
	return;

    DBG(("Win32::OLE->%s()\n", paszMethod[ix]));

    if (items == 0) {
        warn("Win32::OLE->%s must be called as class method", paszMethod[ix]);
	XSRETURN_EMPTY;
    }

    HV *stash = gv_stashsv(ST(0), TRUE);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    switch (ix)
    {
    case 0:
    {
	DWORD dwCoInit = COINIT_MULTITHREADED;
	if (items > 1 && SvOK(ST(1)))
	    dwCoInit = SvIV(ST(1));

	Initialize(PERL_OBJECT_THIS_ gv_stashsv(ST(0), TRUE), dwCoInit);
	break;
    }
    case 1:
    {
	int magic = 0;
	dPERINTERP;

	if (items > 1 && SvOK(ST(1)))
	    magic = SvIV(ST(1));

	Uninitialize(PERL_OBJECT_THIS_ INTERP, magic);
	break;
    }
    case 2:
	SpinMessageLoop();
	break;
    }

    XSRETURN_EMPTY;
}

void
new(...)
PPCODE:
{
    CLSID clsid;
    IDispatch *pDispatch = NULL;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    HRESULT hr;
    STRLEN len;

    if (CallObjectMethod(PERL_OBJECT_THIS_ mark, ax, items, "new"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->new(progid[,destroy])");
	DEBUGBREAK;
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(PERL_OBJECT_THIS_ stash);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(PERL_OBJECT_THIS_ ST(2),
				       "Win32::OLE->new");

    ST(0) = &PL_sv_undef;

    /* normal case: no DCOM */
    if (!SvROK(progid) || SvTYPE(SvRV(progid)) != SVt_PVAV) {
	pBuffer = GetWideChar(PERL_OBJECT_THIS_ SvPV(progid, len), Buffer,
			      OLE_BUF_SIZ, cp);
	if (isalpha(pBuffer[0]))
	    hr = CLSIDFromProgID(pBuffer, &clsid);
	else
	    hr = CLSIDFromString(pBuffer, &clsid);
	ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
	if (SUCCEEDED(hr)) {
	    hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER,
				  IID_IDispatch, (void**)&pDispatch);
	}
	if (!CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
	    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch,
				     destroy);
	    DBG(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
	}
	XSRETURN(1);
    }

    /* DCOM might not exist on Win95 (and does not on NT 3.5) */
    dPERINTERP;
    if (g_pfnCoCreateInstanceEx == NULL) {
	hr = HRESULT_FROM_WIN32(ERROR_SERVICE_DOES_NOT_EXIST);
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	XSRETURN(1);
    }

    /* DCOM spec: ['Servername', 'Program.ID'] */
    AV *av = (AV*)SvRV(progid);
    if (av_len(av) != 1) {
	warn("Win32::OLE->new: for DCOM use ['Machine', 'Prog.Id']");
	XSRETURN(1);
    }
    SV *host = *av_fetch(av, 0, FALSE);
    progid = *av_fetch(av, 1, FALSE);

    /* determine hostname */
    char *pszHost = NULL;
    if (SvPOK(host)) {
	pszHost = SvPV(host, len);
	if (IsLocalMachine(PERL_OBJECT_THIS_ pszHost))
	    pszHost = NULL;
    }

    /* determine CLSID */
    char *pszProgID = SvPV(progid, len);
    pBuffer = GetWideChar(PERL_OBJECT_THIS_ pszProgID, Buffer, OLE_BUF_SIZ, cp);
    if (isalpha(pBuffer[0])) {
	hr = CLSIDFromProgID(pBuffer, &clsid);
	if (FAILED(hr) && pszHost != NULL)
	    hr = CLSIDFromRemoteRegistry(PERL_OBJECT_THIS_ pszHost, pszProgID,
					 &clsid);
    }
    else
        hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    if (FAILED(hr)) {
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	XSRETURN(1);
    }

    /* setup COSERVERINFO & MULTI_QI parameters */
    DWORD clsctx = CLSCTX_REMOTE_SERVER;
    COSERVERINFO ServerInfo;
    OLECHAR ServerName[OLE_BUF_SIZ];
    MULTI_QI multi_qi;

    Zero(&ServerInfo, 1, COSERVERINFO);
    if (pszHost == NULL)
	clsctx = CLSCTX_SERVER;
    else
	ServerInfo.pwszName = GetWideChar(PERL_OBJECT_THIS_ pszHost, ServerName,
					  OLE_BUF_SIZ, cp);

    Zero(&multi_qi, 1, MULTI_QI);
    multi_qi.pIID = &IID_IDispatch;

    /* create instance on remote server */
    hr = g_pfnCoCreateInstanceEx(clsid, NULL, clsctx, &ServerInfo,
				  1, &multi_qi);
    ReleaseBuffer(PERL_OBJECT_THIS_ ServerInfo.pwszName, ServerName);
    if (!CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
	pDispatch = (IDispatch*)multi_qi.pItf;
	ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, destroy);
	DBG(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
    }
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    ReleasePerlObject(PERL_OBJECT_THIS_
		      GetOleObject(PERL_OBJECT_THIS_ self, TRUE));
    XSRETURN_EMPTY;
}

void
Dispatch(self,method,retval,...)
    SV *self
    SV *method
    SV *retval
PPCODE:
{
    char *buffer = "";
    char *ptr;
    unsigned int length, argErr;
    int index, arrayIndex;
    I32 len;
    WINOLEOBJECT *pObj;
    EXCEPINFO excepinfo;
    DISPID dispID = DISPID_VALUE;
    DISPID dispIDParam = DISPID_PROPERTYPUT;
    USHORT wFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
    VARIANT result;
    DISPPARAMS dispParams;
    SV *curitem, *sv;
    HE **rghe = NULL; /* named argument names */

    SV *err = NULL; /* error details */
    HRESULT hr = S_OK;

    ST(0) = &PL_sv_no;
    Zero(&excepinfo, 1, EXCEPINFO);
    VariantInit(&result);

    if (!sv_isobject(self)) {
	warn("Win32::OLE::Dispatch: Cannot be called as class method");
	DEBUGBREAK;
	XSRETURN(1);
    }

    pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL) {
	XSRETURN(1);
    }

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    /* allow [wFlags, 'Method'] instead of 'Method' */
    if (SvROK(method) && (sv = SvRV(method)) &&	SvTYPE(sv) == SVt_PVAV &&
	!SvOBJECT(sv) && av_len((AV*)sv) == 1)
    {
	wFlags = SvIV(*av_fetch((AV*)sv, 0, FALSE));
	method = *av_fetch((AV*)sv, 1, FALSE);
    }

    if (SvPOK(method)) {
	buffer = SvPV(method, length);
	if (length > 0) {
	    hr = GetHashedDispID(PERL_OBJECT_THIS_ pObj, buffer, length, dispID,
				 lcid, cp);
	    if (FAILED(hr)) {
		if (PL_hints & HINT_STRICT_SUBS) {
		    err = newSVpvf(" in GetIDsOfNames of \"%s\"", buffer);
		    ReportOleError(PERL_OBJECT_THIS_ stash, hr, NULL,
				   sv_2mortal(err));
		}
		XSRETURN_EMPTY;
	    }
	}
    }

    DBG(("Dispatch \"%s\"\n", buffer));

    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = items - 3;

    /* last arg is ref to a non-object-hash => named arguments */
    curitem = ST(items-1);
    if (SvROK(curitem) && (sv = SvRV(curitem)) &&
	SvTYPE(sv) == SVt_PVHV && !SvOBJECT(sv))
    {
	if (wFlags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF)) {
	    warn("Win32::OLE->Dispatch: named arguments not supported "
		 "for PROPERTYPUT");
	    DEBUGBREAK;
	    XSRETURN_EMPTY;
	}

	OLECHAR **rgszNames;
	DISPID  *rgdispids;
	HV      *hv = (HV*)sv;

	dispParams.cNamedArgs = HvKEYS(hv);
	dispParams.cArgs += dispParams.cNamedArgs - 1;

	New(0, rghe, dispParams.cNamedArgs, HE *);
	New(0, dispParams.rgdispidNamedArgs, dispParams.cNamedArgs, DISPID);
	New(0, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	for (index = 0 ; index < dispParams.cArgs ; ++index)
	    VariantInit(&dispParams.rgvarg[index]);

	New(0, rgszNames, 1+dispParams.cNamedArgs, OLECHAR *);
	New(0, rgdispids, 1+dispParams.cNamedArgs, DISPID);

	rgszNames[0] = AllocOleString(PERL_OBJECT_THIS_ buffer, length, cp);
	hv_iterinit(hv);
	for (index = 0; index < dispParams.cNamedArgs; ++index) {
	    rghe[index] = hv_iternext(hv);
	    char *pszName = hv_iterkey(rghe[index], &len);
	    rgszNames[1+index] = AllocOleString(PERL_OBJECT_THIS_ pszName,
						len, cp);
	}

	hr = pObj->pDispatch->GetIDsOfNames(IID_NULL, rgszNames,
			      1+dispParams.cNamedArgs, lcid, rgdispids);

	if (SUCCEEDED(hr)) {
	    for (index = 0; index < dispParams.cNamedArgs; ++index) {
		dispParams.rgdispidNamedArgs[index] = rgdispids[index+1];
		hr = SetVariantFromSV(PERL_OBJECT_THIS_
				      hv_iterval(hv, rghe[index]),
				      &dispParams.rgvarg[index], cp);
		if (FAILED(hr))
		    break;
	    }
	}
	else {
	    unsigned int cErrors = 0;
	    unsigned int error = 0;

	    for (index = 1 ; index <= dispParams.cNamedArgs ; ++index)
		if (rgdispids[index] == DISPID_UNKNOWN)
		   ++cErrors;

	    err = sv_2mortal(newSVpv("",0));
	    for (index = 1 ; index <= dispParams.cNamedArgs ; ++index)
		if (rgdispids[index] == DISPID_UNKNOWN) {
		    if (error++ > 0)
			sv_catpv(err, error == cErrors ? " and " : ", ");
		    sv_catpvf(err, "\"%s\"", hv_iterkey(rghe[index-1], &len));
		}
	    sv_catpvf(err, " in GetIDsOfNames for \"%s\"", buffer);
	}

	for (index = 0; index <= dispParams.cNamedArgs; ++index)
	    SysFreeString(rgszNames[index]);
	Safefree(rgszNames);
	Safefree(rgdispids);

	if (FAILED(hr))
	    goto Cleanup;

	--items;
    }

    if (dispParams.cArgs > dispParams.cNamedArgs) {
	if (dispParams.rgvarg == NULL) {
	    New(0, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	    for (index = 0 ; index < dispParams.cArgs ; ++index)
		VariantInit(&dispParams.rgvarg[index]);
	}

	for(index = dispParams.cNamedArgs; index < dispParams.cArgs; ++index) {
	    hr = SetVariantFromSV(PERL_OBJECT_THIS_
				  ST(items-1-(index-dispParams.cNamedArgs)),
				  &dispParams.rgvarg[index], cp);
	    if (FAILED(hr))
		goto Cleanup;
	}
    }

    if (wFlags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF)) {
	Safefree(dispParams.rgdispidNamedArgs);
	dispParams.rgdispidNamedArgs = &dispIDParam;
	dispParams.cNamedArgs = 1;
    }

    hr = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid, wFlags,
				  &dispParams, &result, &excepinfo, &argErr);

    if (FAILED(hr)) {
	/* mega kludge. if a method in WORD is called and we ask
	 * for a result when one is not returned then
	 * hResult == DISP_E_EXCEPTION. this only happens on
	 * functions whose DISPID > 0x8000 */

	if (hr == DISP_E_EXCEPTION && dispID > 0x8000) {
	    Zero(&excepinfo, 1, EXCEPINFO);
	    hr = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid, wFlags,
				  &dispParams, NULL, &excepinfo, &argErr);
	}
    }

    if (SUCCEEDED(hr)) {
	if (sv_isobject(retval) && sv_derived_from(retval, szWINOLEVARIANT)) {
	    WINOLEVARIANTOBJECT *pVarObj =
		GetOleVariantObject(PERL_OBJECT_THIS_ retval);

	    if (pVarObj != NULL) {
		VariantClear(&pVarObj->byref);
		VariantClear(&pVarObj->variant);
		VariantCopy(&pVarObj->variant, &result);
		ST(0) = &PL_sv_yes;
	    }
	}
	else {
	    hr = SetSVFromVariantEx(PERL_OBJECT_THIS_ &result, retval, stash);
	    ST(0) = &PL_sv_yes;
	}
    }
    else {
	/* use more specific error code from exception when available */
	if (hr == DISP_E_EXCEPTION && FAILED(excepinfo.scode))
	    hr = excepinfo.scode;

	char *pszDelim = "";
	err = sv_newmortal();
	sv_setpvf(err, "in ");

	if (wFlags&DISPATCH_METHOD) {
	    sv_catpv(err, "METHOD");
	    pszDelim = "/";
	}
	if (wFlags&DISPATCH_PROPERTYGET) {
	    sv_catpvf(err, "%sPROPERTYGET", pszDelim);
	    pszDelim = "/";
	}
	if (wFlags&DISPATCH_PROPERTYPUT) {
	    sv_catpvf(err, "%sPROPERTYPUT", pszDelim);
	    pszDelim = "/";
	}
	if (wFlags&DISPATCH_PROPERTYPUTREF)
	    sv_catpvf(err, "%sPROPERTYPUTREF", pszDelim);

	sv_catpvf(err, " \"%s\"", buffer);

	if (hr == DISP_E_TYPEMISMATCH || hr == DISP_E_PARAMNOTFOUND) {
	    if (argErr < dispParams.cNamedArgs)
		sv_catpvf(err, " argument \"%s\"",
			  hv_iterkey(rghe[argErr], &len));
	    else
		sv_catpvf(err, " argument %d", 1 + dispParams.cArgs - argErr);
	}
    }

 Cleanup:
    VariantClear(&result);
    if (dispParams.cArgs != 0 && dispParams.rgvarg != NULL) {
	for(index = 0; index < dispParams.cArgs; ++index)
	    VariantClear(&dispParams.rgvarg[index]);
	Safefree(dispParams.rgvarg);
    }
    Safefree(rghe);
    if (dispParams.rgdispidNamedArgs != &dispIDParam)
	Safefree(dispParams.rgdispidNamedArgs);

    CheckOleError(PERL_OBJECT_THIS_ stash, hr, &excepinfo, err);

    XSRETURN(1);
}

void
GetActiveObject(...)
PPCODE:
{
    CLSID clsid;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    unsigned int length;
    char *buffer;
    HRESULT hr;
    IUnknown *pUnknown;
    IDispatch *pDispatch;

    if (CallObjectMethod(PERL_OBJECT_THIS_ mark, ax, items, "GetActiveObject"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->GetActiveObject(progid[,destroy])");
	DEBUGBREAK;
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(PERL_OBJECT_THIS_ stash);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(PERL_OBJECT_THIS_ ST(2),
				       "Win32::OLE->GetActiveObject");

    buffer = SvPV(progid, length);
    pBuffer = GetWideChar(PERL_OBJECT_THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    if (isalpha(pBuffer[0]))
        hr = CLSIDFromProgID(pBuffer, &clsid);
    else
        hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = GetActiveObject(clsid, 0, &pUnknown);
    /* Don't call CheckOleError! Return "undef" for "Server not running" */
    if (FAILED(hr))
	XSRETURN_EMPTY;

    hr = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, destroy);
    DBG(("Win32::OLE::GetActiveObject |%lx| |%lx|\n", ST(0), pDispatch));
    XSRETURN(1);
}

void
GetObject(...)
PPCODE:
{
    IBindCtx *pBindCtx;
    IMoniker *pMoniker;
    IDispatch *pDispatch;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    char *buffer;
    ULONG ulEaten;
    HRESULT hr;
    STRLEN len;

    if (CallObjectMethod(PERL_OBJECT_THIS_ mark, ax, items, "GetObject"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->GetObject(pathname[,destroy])");
	DEBUGBREAK;
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *pathname = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(PERL_OBJECT_THIS_ stash);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(PERL_OBJECT_THIS_ ST(2),
				       "Win32::OLE->GetObject");

    hr = CreateBindCtx(0, &pBindCtx);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    buffer = SvPV(pathname, len);
    pBuffer = GetWideChar(PERL_OBJECT_THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    hr = MkParseDisplayName(pBindCtx, pBuffer, &ulEaten, &pMoniker);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    if (FAILED(hr)) {
	pBindCtx->Release();
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "after character %lu in \"%s\"", ulEaten, buffer);
	ReportOleError(PERL_OBJECT_THIS_ stash, hr, NULL, sv);
	XSRETURN_EMPTY;
    }

    hr = pMoniker->BindToObject(pBindCtx, NULL, IID_IDispatch,
				 (void**)&pDispatch);
    pBindCtx->Release();
    pMoniker->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, destroy);
    XSRETURN(1);
}

void
GetTypeInfo(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    HV *stash = gv_stashsv(self, TRUE);
    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);

    SetLastOleError(PERL_OBJECT_THIS_ stash);
    HRESULT hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(PERL_OBJECT_THIS_ pTypeInfo,
					    pTypeAttr));
    XSRETURN(1);
}

void
QueryInterface(self,itf)
    SV *self
    SV *itf
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    IID iid;
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;

    // XXX support GUIDs in addition to names too
    STRLEN len;
    char *szItf = SvPV(itf, len);

    HV *stash = SvSTASH(pObj->self);
    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    SetLastOleError(PERL_OBJECT_THIS_ stash);

    // Determine containing type library
    HRESULT hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    unsigned int index;
    hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
    pTypeInfo->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    // Walk through all type definitions in the library
    BOOL bFound = FALSE;
    unsigned int count = pTypeLib->GetTypeInfoCount();
    hr = S_OK;
    for (index = 0; index < count; ++index) {
	TYPEATTR *pTypeAttr;

	hr = pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (FAILED(hr))
	    break;

	hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    break;
	}

	// Look into all COCLASSes
	if (pTypeAttr->typekind == TKIND_COCLASS) {

	    // Walk through all implemented types
	    for (unsigned int type=0; type < pTypeAttr->cImplTypes; ++type) {
		HREFTYPE RefType;
		ITypeInfo *pImplTypeInfo;

		hr = pTypeInfo->GetRefTypeOfImplType(type, &RefType);
		if (FAILED(hr))
		    break;

		hr = pTypeInfo->GetRefTypeInfo(RefType, &pImplTypeInfo);
		if (FAILED(hr))
		    break;

		BSTR bstr;
		hr = pImplTypeInfo->GetDocumentation(-1, &bstr, NULL,
						     NULL, NULL);
		if (CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
		    pImplTypeInfo->Release();
		    break;
		}

		char szStr[OLE_BUF_SIZ];
		char *pszStr = GetMultiByte(PERL_OBJECT_THIS_ bstr, szStr,
					    sizeof(szStr), cp);
		if (strEQ(szItf, pszStr)) {
		    TYPEATTR *pImplTypeAttr;

		    hr = pImplTypeInfo->GetTypeAttr(&pImplTypeAttr);
		    if (SUCCEEDED(hr)) {
			bFound = TRUE;
			iid = pImplTypeAttr->guid;
			pImplTypeInfo->ReleaseTypeAttr(pImplTypeAttr);
		    }
		}

		ReleaseBuffer(PERL_OBJECT_THIS_ pszStr, szStr);
		pImplTypeInfo->Release();
		if (bFound || FAILED(hr))
		    break;
	    }
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeInfo->Release();
	if (bFound || FAILED(hr))
	    break;
    }

    pTypeLib->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    if (!bFound) {
	warn("Win32::OLE->QueryInterface: Interface '%s' not found", szItf);
        XSRETURN_EMPTY;
    }

    IUnknown *pUnknown;
    IDispatch *pDispatch;

    if (0) {
	OLECHAR wszGUID[80];
	int len = StringFromGUID2(iid, wszGUID, sizeof(wszGUID)/sizeof(OLECHAR));
	char szStr[OLE_BUF_SIZ];
	char *pszStr = GetMultiByte(PERL_OBJECT_THIS_ wszGUID, szStr,
				    sizeof(szStr), cp);
	warn("iid is %s", pszStr);
	ReleaseBuffer(PERL_OBJECT_THIS_ pszStr, szStr);
    }

    hr = pObj->pDispatch->QueryInterface(iid, (void**)&pUnknown);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    hr = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, NULL);
    DBG(("Win32::OLE::QueryInterface |%lx| |%lx|\n", ST(0), pDispatch));
    XSRETURN(1);
}

void
QueryObjectType(...)
PPCODE:
{
    if (CallObjectMethod(PERL_OBJECT_THIS_ mark, ax, items, "QueryObjectType"))
	return;

    if (items != 2) {
	warn("Usage: Win32::OLE->QueryObjectType(object)");
	DEBUGBREAK;
	XSRETURN_EMPTY;
    }

    SV *object = ST(1);

    if (!sv_isobject(object) || !sv_derived_from(object, szWINOLE)) {
	warn("Win32::OLE->QueryObjectType: object is not a Win32::OLE object");
	XSRETURN_EMPTY;
    }

    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ object);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int count;
    BSTR bstr;

    HRESULT hr = pObj->pDispatch->GetTypeInfoCount(&count);
    if (FAILED(hr) || count == 0)
	XSRETURN_EMPTY;

    HV *stash = gv_stashsv(ST(0), TRUE);
    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    SetLastOleError(PERL_OBJECT_THIS_ stash);
    hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    /* Return ('TypeLib Name', 'Class Name') in array context */
    if (GIMME_V == G_ARRAY) {
	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	    XSRETURN_EMPTY;
	}

	hr = pTypeLib->GetDocumentation(-1, &bstr, NULL, NULL, NULL);
	pTypeLib->Release();
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	    XSRETURN_EMPTY;
	}

	PUSHs(sv_2mortal(sv_setwide(PERL_OBJECT_THIS_ NULL, bstr, cp)));
	SysFreeString(bstr);
    }

    hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstr, NULL, NULL, NULL);
    pTypeInfo->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_EMPTY;

    PUSHs(sv_2mortal(sv_setwide(PERL_OBJECT_THIS_ NULL, bstr, cp)));
    SysFreeString(bstr);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Tie

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self, TRUE);
    if (pObj != NULL) {
	DBG(("Win32::OLE::Tie::DESTROY |%lx| |%lx|\n", pObj, pObj->pDispatch));
	RemoveFromObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER *)pObj);
	Safefree(pObj);
    }
    XSRETURN_EMPTY;
}

void
Fetch(self,key,def)
    SV *self
    SV *key
    SV *def
PPCODE:
{
    char *buffer;
    unsigned int length;
    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    VARIANTARG propName;
    DISPID dispID = DISPID_VALUE;
    HRESULT hr;

    buffer = SvPV(key, length);
    if (strEQ(buffer, PERL_OLE_ID)) {
	ST(0) = *hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
	XSRETURN(1);
    }

    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    ST(0) = &PL_sv_undef;
    VariantInit(&result);
    VariantInit(&propName);

    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    dispParams.cArgs = 0;
    dispParams.rgvarg = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    hr = GetHashedDispID(PERL_OBJECT_THIS_ pObj, buffer, length, dispID,
			 lcid, cp);
    if (FAILED(hr)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr, NULL, sv_2mortal(err));
	    XSRETURN(1);
	}

	/* default method call: $self->{Key} ---> $self->Item('Key') */
	V_VT(&propName) = VT_BSTR;
	V_BSTR(&propName) = AllocOleString(PERL_OBJECT_THIS_ buffer, length,
					   cp);
	dispParams.cArgs = 1;
	dispParams.rgvarg = &propName;
    }

    Zero(&excepinfo, 1, EXCEPINFO);

    hr = pObj->pDispatch->Invoke(dispID, IID_NULL,
		    lcid, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
		    &dispParams, &result, &excepinfo, &argErr);

    VariantClear(&propName);

    if (FAILED(hr)) {
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "in METHOD/PROPERTYGET \"%s\"", buffer);
	VariantClear(&result);
	ReportOleError(PERL_OBJECT_THIS_ stash, hr, &excepinfo, sv);
    }
    else {
	ST(0) = sv_newmortal();
	hr = SetSVFromVariantEx(PERL_OBJECT_THIS_ &result, ST(0), stash);
	VariantClear(&result);
	CheckOleError(PERL_OBJECT_THIS_ stash, hr);
    }

    XSRETURN(1);
}

void
Store(self,key,value,def)
    SV *self
    SV *key
    SV *value
    SV *def
PPCODE:
{
    unsigned int length, argErr;
    char *buffer;
    int index;
    HRESULT hr;
    EXCEPINFO excepinfo;
    DISPID dispID = DISPID_VALUE;
    DISPID dispIDParam = DISPID_PROPERTYPUT;
    DISPPARAMS dispParams;
    VARIANTARG propertyValue[2];
    SV *err = NULL;

    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    dispParams.rgdispidNamedArgs = &dispIDParam;
    dispParams.rgvarg = propertyValue;
    dispParams.cNamedArgs = 1;
    dispParams.cArgs = 1;

    VariantInit(&propertyValue[0]);
    VariantInit(&propertyValue[1]);
    Zero(&excepinfo, 1, EXCEPINFO);

    buffer = SvPV(key, length);
    hr = GetHashedDispID(PERL_OBJECT_THIS_ pObj, buffer, length, dispID, lcid,
			 cp);
    if (FAILED(hr)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr, NULL, sv_2mortal(err));
	    XSRETURN_EMPTY;
	}

	dispParams.cArgs = 2;
	V_VT(&propertyValue[1]) = VT_BSTR;
	V_BSTR(&propertyValue[1]) = AllocOleString(PERL_OBJECT_THIS_ buffer,
						   length, cp);
    }

    hr = SetVariantFromSV(PERL_OBJECT_THIS_ value, &propertyValue[0], cp);
    if (SUCCEEDED(hr)) {
	USHORT wFlags = DISPATCH_PROPERTYPUT;

	/* objects are passed by reference */
	VARTYPE vt = V_VT(&propertyValue[0]);
	if (vt == VT_DISPATCH || vt == VT_UNKNOWN)
	    wFlags = DISPATCH_PROPERTYPUTREF;

	hr = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid, wFlags,
				      &dispParams, NULL, &excepinfo, &argErr);
	if (FAILED(hr)) {
	    err = sv_newmortal();
	    sv_setpvf(err, "in PROPERTYPUT%s \"%s\"",
		      (wFlags == DISPATCH_PROPERTYPUTREF ? "REF" : ""), buffer);
	}
    }

    for(index = 0; index < dispParams.cArgs; ++index)
	VariantClear(&propertyValue[index]);

    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr, &excepinfo, err))
	XSRETURN_EMPTY;

    XSRETURN_YES;
}

void
FIRSTKEY(self,...)
    SV *self
ALIAS:
    NEXTKEY   = 1
    FIRSTENUM = 2
    NEXTENUM  = 3
PPCODE:
{
    /* NEXTKEY has an additional "lastkey" arg, which is not needed here */
    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ self);
    char *paszMethod[] = {"FIRSTKEY", "NEXTKEY", "FIRSTENUM", "NEXTENUM"};

    DBG(("%s called, pObj=%p\n", paszMethod[ix], pObj));
    if (pObj == NULL)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    switch (ix)
    {
    case 0: /* FIRSTKEY */
	FetchTypeInfo(PERL_OBJECT_THIS_ pObj);
	pObj->PropIndex = 0;
    case 1: /* NEXTKEY */
	ST(0) = NextPropertyName(PERL_OBJECT_THIS_ pObj);
	break;

    case 2: /* FIRSTENUM */
	if (pObj->pEnum != NULL)
	    pObj->pEnum->Release();
	pObj->pEnum = CreateEnumVARIANT(PERL_OBJECT_THIS_ pObj);
    case 3: /* NEXTENUM */
	ST(0) = NextEnumElement(PERL_OBJECT_THIS_ pObj->pEnum, stash);
	if (!SvOK(ST(0))) {
	    pObj->pEnum->Release();
	    pObj->pEnum = NULL;
	}
	break;
    }

    if (!SvIMMORTAL(ST(0)))
	sv_2mortal(ST(0));

    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Const

void
_Load(classid,major,minor,locale,typelib,codepage,caller)
    SV *classid
    IV major
    IV minor
    SV *locale
    SV *typelib
    SV *codepage
    SV *caller
PPCODE:
{
    ITypeLib *pTypeLib;
    CLSID clsid;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    HRESULT hr;
    LCID lcid = lcidDefault;
    UINT cp = cpDefault;
    HV *stash = gv_stashpv(szWINOLE, TRUE);
    HV *hv;
    unsigned int count;

    Initialize(PERL_OBJECT_THIS_ stash);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    if (SvIOK(locale))
	lcid = SvIV(locale);

    if (SvIOK(codepage))
	cp = SvIV(codepage);

    if (sv_derived_from(classid, szWINOLE)) {
	/* Get containing typelib from IDispatch interface */
	ITypeInfo *pTypeInfo;
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ classid);
	if (pObj == NULL)
	    XSRETURN_EMPTY;

	stash = SvSTASH(pObj->self);
	hr = pObj->pDispatch->GetTypeInfoCount(&count);
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) || count == 0)
	    XSRETURN_EMPTY;

	lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			   lcidDefault);
	cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

	hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_EMPTY;

	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	pTypeInfo->Release();
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_EMPTY;
    }
    else {
	/* try to load registered typelib by classid, version and lcid */
	STRLEN len;
	char *pszBuffer = SvPV(classid, len);
	pBuffer = GetWideChar(PERL_OBJECT_THIS_ pszBuffer,
			      Buffer, OLE_BUF_SIZ, cp);
	hr = CLSIDFromString(pBuffer, &clsid);
	ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);

	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_EMPTY;

	hr = LoadRegTypeLib(clsid, major, minor, lcid, &pTypeLib);
	if (FAILED(hr) && SvPOK(typelib)) {
	    /* typelib not registerd, try to read from file "typelib" */
	    pszBuffer = SvPV(typelib, len);
	    pBuffer = GetWideChar(PERL_OBJECT_THIS_ pszBuffer,
				  Buffer, OLE_BUF_SIZ, cp);
	    hr = LoadTypeLib(pBuffer, &pTypeLib);
	    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
	}
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_EMPTY;
    }

    if (SvOK(caller)) {
	/* we'll define inlineable functions returning a const */
        hv = gv_stashsv(caller, TRUE);
	ST(0) = &PL_sv_undef;
    }
    else {
	/* we'll return ref to hash with constant name => value pairs */
	hv = newHV();
        ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    }

    /* loop through all objects in type lib */
    count = pTypeLib->GetTypeInfoCount();
    for (int index=0 ; index < count ; ++index) {
	ITypeInfo *pTypeInfo;
	TYPEATTR  *pTypeAttr;

	hr = pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    continue;

	hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	    continue;
	}

	for (int iVar=0 ; iVar < pTypeAttr->cVars ; ++iVar) {
	    VARDESC *pVarDesc;

	    hr = pTypeInfo->GetVarDesc(iVar, &pVarDesc);
	    /* XXX LEAK alert */
	    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	        continue;

	    if (pVarDesc->varkind == VAR_CONST &&
		!(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
					 VARFLAG_FRESTRICTED |
					 VARFLAG_FNONBROWSABLE)))
	    {
		unsigned int cName;
		BSTR bstr;
		char szName[64];

		hr = pTypeInfo->GetNames(pVarDesc->memid, &bstr, 1, &cName);
		if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) ||
		    cName == 0 || bstr == NULL)
		{
		    continue;
		}

		char *pszName = GetMultiByte(PERL_OBJECT_THIS_ bstr,
					     szName, sizeof(szName), cp);
		SV *sv = newSVpv("",0);
		/* XXX LEAK alert */
		hr = SetSVFromVariantEx(PERL_OBJECT_THIS_ pVarDesc->lpvarValue,
					sv, stash);
		if (!CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
		    if (SvOK(caller)) {
			/* XXX check for valid symbol name */
			newCONSTSUB(hv, pszName, sv);
		    }
		    else
		        hv_store(hv, pszName, strlen(pszName), sv, 0);
		}
		SysFreeString(bstr);
		ReleaseBuffer(PERL_OBJECT_THIS_ pszName, szName);
	    }
	    pTypeInfo->ReleaseVarDesc(pVarDesc);
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeInfo->Release();
    }

    pTypeLib->Release();

    XSRETURN(1);
}

void
_Typelibs(self)
    SV *self
PPCODE:
{
    HKEY hKeyTypelib;
    FILETIME ft;
    LONG err;

    err = RegOpenKeyEx(HKEY_CLASSES_ROOT, "Typelib", 0, KEY_READ, &hKeyTypelib);
    if (err != ERROR_SUCCESS) {
	warn("Cannot access HKEY_CLASSES_ROOT\\Typelib");
	XSRETURN_EMPTY;
    }

    AV *res = newAV();

    // Enumerate all Clsids
    for (DWORD dwClsid=0;; ++dwClsid) {
	char szClsid[100];
	DWORD cbClsid = sizeof(szClsid);
	err = RegEnumKeyEx(hKeyTypelib, dwClsid, szClsid, &cbClsid,
			   NULL, NULL, NULL, &ft);
	if (err != ERROR_SUCCESS)
	    break;

	HKEY hKeyClsid;
	err = RegOpenKeyEx(hKeyTypelib, szClsid, 0, KEY_READ, &hKeyClsid);
	if (err != ERROR_SUCCESS)
	    continue;

	// Enumerate versions for current clsid
	for (DWORD dwVersion=0;; ++dwVersion) {
	    char szVersion[10];
	    DWORD cbVersion = sizeof(szVersion);
	    err = RegEnumKeyEx(hKeyClsid, dwVersion, szVersion, &cbVersion,
			       NULL, NULL, NULL, &ft);
	    if (err != ERROR_SUCCESS)
		break;

	    HKEY hKeyVersion;
	    err = RegOpenKeyEx(hKeyClsid, szVersion, 0, KEY_READ, &hKeyVersion);
	    if (err != ERROR_SUCCESS)
		continue;

	    char szTitle[300];
	    LONG cbTitle = sizeof(szTitle);
	    err = RegQueryValue(hKeyVersion, NULL, szTitle, &cbTitle);
	    if (err != ERROR_SUCCESS || cbTitle <= 1)
		continue;

	    // Enumerate languages
	    for (DWORD dwLangid=0;; ++dwLangid) {
		char szLangid[10];
		DWORD cbLangid = sizeof(szLangid);
		err = RegEnumKeyEx(hKeyVersion, dwLangid, szLangid, &cbLangid,
				   NULL, NULL, NULL, &ft);
		if (err != ERROR_SUCCESS)
		    break;

		// Language ids must be strictly numeric
		char *psz=szLangid;
		while (isDIGIT(*psz))
		    ++psz;
		if (*psz)
		    continue;

		HKEY hKeyLangid;
		err = RegOpenKeyEx(hKeyVersion, szLangid, 0, KEY_READ, &hKeyLangid);
		if (err != ERROR_SUCCESS)
		    continue;

		// Retrieve filename of type library
		char szFile[MAX_PATH+1];
		LONG cbFile = sizeof(szFile);
		err = RegQueryValue(hKeyLangid, "win32", szFile, &cbFile);
		if (err == ERROR_SUCCESS && cbFile > 1) {
		    AV *av = newAV();
		    av_push(av, newSVpv(szClsid, cbClsid));
		    av_push(av, newSVpv(szTitle, cbTitle));
		    av_push(av, newSVpv(szVersion, cbVersion));
		    av_push(av, newSVpv(szLangid, cbLangid));
		    av_push(av, newSVpv(szFile, cbFile-1));
		    av_push(res, newRV_noinc((SV*)av));
		}

		RegCloseKey(hKeyLangid);
	    }
	    RegCloseKey(hKeyVersion);
	}
	RegCloseKey(hKeyClsid);
    }
    RegCloseKey(hKeyTypelib);

    ST(0) = sv_2mortal(newRV_noinc((SV*)res));
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Enum

void
new(self,object)
    SV *self
    SV *object
ALIAS:
    Clone = 1
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj;
    New(0, pEnumObj, 1, WINOLEENUMOBJECT);

    if (ix == 0) { /* new */
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ object);
	if (pObj != NULL) {
	    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ object);
	    SetLastOleError(PERL_OBJECT_THIS_ olestash);
	    pEnumObj->pEnum = CreateEnumVARIANT(PERL_OBJECT_THIS_ pObj);
	}
    }
    else { /* Clone */
	WINOLEENUMOBJECT *pOriginal = GetOleEnumObject(PERL_OBJECT_THIS_ self);
	if (pOriginal != NULL) {
	    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
	    SetLastOleError(PERL_OBJECT_THIS_ olestash);

	    HRESULT hr = pOriginal->pEnum->Clone(&pEnumObj->pEnum);
	    CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
	}
    }

    if (pEnumObj->pEnum == NULL) {
	Safefree(pEnumObj);
	XSRETURN_EMPTY;
    }

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pEnumObj,
		     WINOLEENUM_MAGIC);

    SV *sv = newSViv((IV)pEnumObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv),
				GetStash(PERL_OBJECT_THIS_ self)));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(PERL_OBJECT_THIS_ self, TRUE);
    if (pEnumObj != NULL) {
	RemoveFromObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pEnumObj);
	if (pEnumObj->pEnum != NULL)
	    pEnumObj->pEnum->Release();
	Safefree(pEnumObj);
    }
    XSRETURN_EMPTY;
}

void
All(self,...)
    SV *self
ALIAS:
    Next = 1
PPCODE:
{
    int count = 1;
    if (ix == 0) { /* All */
	/* my @list = Win32::OLE::Enum->All($Excel->Workbooks); */
	if (!sv_isobject(self) && items > 1) {
	    /* $self = $self->new(shift); */
	    SV *obj = ST(1);
	    PUSHMARK(sp);
	    PUSHs(self);
	    PUSHs(obj);
	    PUTBACK;
	    items = perl_call_method("new", G_SCALAR);
	    SPAGAIN;
	    if (items == 1)
		self = POPs;
	    PUTBACK;
	}
    }
    else { /* Next */
	if (items > 1)
	    count = SvIV(ST(1));
	if (count < 1) {
	    warn(MY_VERSION ": Win32::OLE::Enum::Next: invalid Count %ld",
		 count);
	    DEBUGBREAK;
	    count = 1;
	}
    }

    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(PERL_OBJECT_THIS_ self);
    if (pEnumObj == NULL)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    SetLastOleError(PERL_OBJECT_THIS_ olestash);

    SV *sv = NULL;
    while (ix == 0 || count-- > 0) {
	sv = NextEnumElement(PERL_OBJECT_THIS_ pEnumObj->pEnum, olestash);
	if (!SvOK(sv))
	    break;
	if (!SvIMMORTAL(sv))
	    sv_2mortal(sv);
	if (GIMME_V == G_ARRAY)
	    XPUSHs(sv);
    }

    if (GIMME_V == G_SCALAR && sv != NULL && SvOK(sv))
	XPUSHs(sv);
}

void
Reset(self)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(PERL_OBJECT_THIS_ self);
    if (pEnumObj == NULL)
	XSRETURN_NO;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    SetLastOleError(PERL_OBJECT_THIS_ olestash);

    HRESULT hr = pEnumObj->pEnum->Reset();
    CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
    ST(0) = boolSV(hr == S_OK);
    XSRETURN(1);
}

void
Skip(self,...)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(PERL_OBJECT_THIS_ self);
    if (pEnumObj == NULL)
	XSRETURN_NO;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    SetLastOleError(PERL_OBJECT_THIS_ olestash);
    int count = (items > 1) ? SvIV(ST(1)) : 1;
    HRESULT hr = pEnumObj->pEnum->Skip(count);
    CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
    ST(0) = boolSV(hr == S_OK);
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Variant

void
new(self,...)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj;
    VARTYPE vt = items < 2 ? VT_EMPTY : SvIV(ST(1));
    SV *data = items < 3 ? NULL : ST(2);

    // XXX Initialize should be superfluous here
    // Initialize();
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    SetLastOleError(PERL_OBJECT_THIS_ olestash);

    VARTYPE vt_base = vt & ~VT_BYREF & ~VT_ARRAY;
    if (data == NULL && vt_base != VT_NULL && vt_base != VT_EMPTY) {
	warn(MY_VERSION ": Win32::OLE::Variant->new(vt, data): data may be"
	     " omitted only for VT_NULL or VT_EMPTY");
	XSRETURN_EMPTY;
    }

    Newz(0, pVarObj, 1, WINOLEVARIANTOBJECT);
    VARIANT *pVariant = &pVarObj->variant;
    VariantInit(pVariant);
    VariantInit(&pVarObj->byref);

    V_VT(pVariant) = vt;
    if (vt & VT_BYREF) {
	if ((vt & ~VT_BYREF) == VT_VARIANT)
	    V_VARIANTREF(pVariant) = &pVarObj->byref;
	else
	    V_BYREF(pVariant) = &V_UI1(&pVarObj->byref);
    }

    if (vt & VT_ARRAY) {
	UINT cDims = items - 2;
	SAFEARRAYBOUND *rgsabound;

	if (cDims == 0) {
	    warn(MY_VERSION ": Win32::OLE::Variant->new() VT_ARRAY but "
		 "no array dimensions specified");
	    Safefree(pVarObj);
	    XSRETURN_EMPTY;
	}

	Newz(0, rgsabound, cDims, SAFEARRAYBOUND);
	for (int iDim=0; iDim < cDims; ++iDim) {
	    SV *sv = ST(2+iDim);

	    if (SvROK(sv) && SvTYPE(SvRV(sv)) == SVt_PVAV) {
		AV *av = (AV*)SvRV(sv);
		SV **elt = av_fetch(av, 0, FALSE);
		if (elt != NULL)
		    rgsabound[iDim].lLbound = SvIV(*elt);
		rgsabound[iDim].cElements = 1;
		elt = av_fetch(av, 1, FALSE);
		if (elt != NULL)
		    rgsabound[iDim].cElements +=
			SvIV(*elt) - rgsabound[iDim].lLbound;
	    }
	    else
		rgsabound[iDim].cElements = SvIV(sv);
	}

	SAFEARRAY *psa = SafeArrayCreate(vt_base, cDims, rgsabound);
	Safefree(rgsabound);
	if (psa == NULL) {
	    /* XXX No HRESULT value available */
	    warn(MY_VERSION ": Win32::OLE::Variant->new() couldnot "
		 "allocate SafeArray");
	    Safefree(pVarObj);
	    XSRETURN_EMPTY;
	}

	if (vt & VT_BYREF)
	    *V_ARRAYREF(pVariant) = psa;
	else
	    V_ARRAY(pVariant) = psa;
    }
    else if (vt == VT_UI1 && SvPOK(data)) {
	/* Special case: VT_UI1 with string implies VT_ARRAY */
	unsigned char* pDest;
	STRLEN len;
	char *ptr = SvPV(data, len);
	V_ARRAY(pVariant) = SafeArrayCreateVector(VT_UI1, 0, len);
	if (V_ARRAY(pVariant) != NULL) {
	    V_VT(pVariant) = VT_UI1 | VT_ARRAY;
	    HRESULT hr = SafeArrayAccessData(V_ARRAY(pVariant), (void**)&pDest);
	    if (FAILED(hr)) {
		VariantClear(pVariant);
		ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	    }
	    else {
		memcpy(pDest, ptr, len);
		SafeArrayUnaccessData(V_ARRAY(pVariant));
	    }
	}
    }
    else {
	AssignVariantFromSV(PERL_OBJECT_THIS_ data, pVariant, olestash);
    }

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pVarObj,
		     WINOLEVARIANT_MAGIC);

    HV *stash = GetStash(PERL_OBJECT_THIS_ self);
    SV *sv = newSViv((IV)pVarObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);
    if (pVarObj != NULL) {
	RemoveFromObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pVarObj);
	VariantClear(&pVarObj->byref);
	VariantClear(&pVarObj->variant);
	Safefree(pVarObj);
    }

    XSRETURN_EMPTY;
}

void
As(self,type)
    SV *self
    IV type
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj != NULL) {
	HRESULT hr;
	VARIANT variant;
	HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ olestash, LCID_NAME, LCID_LEN,
				lcidDefault);

	SetLastOleError(PERL_OBJECT_THIS_ olestash);
	VariantInit(&variant);
	hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, type);
	if (SUCCEEDED(hr)) {
	    ST(0) = sv_newmortal();
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ &variant, ST(0), olestash);
	}
	VariantClear(&variant);
	CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
    }
    XSRETURN(1);
}

void
ChangeType(self,type)
    SV *self
    IV type
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);
    HRESULT hr = E_INVALIDARG;

    if (pVarObj != NULL) {
	HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ olestash, LCID_NAME, LCID_LEN,
				lcidDefault);

	SetLastOleError(PERL_OBJECT_THIS_ olestash);
	/* XXX: Does it work with VT_BYREF? */
	hr = VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant,
				  lcid, 0, type);
	CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
    }

    if (FAILED(hr))
	ST(0) = &PL_sv_undef;

    XSRETURN(1);
}

void
Copy(self,...)
    SV *self
ALIAS:
    _Clone = 1
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);
    if (pVarObj == NULL)
	XSRETURN_EMPTY;

    HRESULT hr;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);

    VARIANT *pSource = &pVarObj->variant;
    VARIANT variant, byref;
    VariantInit(&variant);
    VariantInit(&byref);

    /* Copy(DIM) makes a copy of a SAFEARRAY element */
    if (items > 1) {
	if (ix != 0) {
	    warn(MY_VERSION ": Win32::OLE::Variant->_Clone doesn't support "
		 "array elements");
	    XSRETURN_EMPTY;
	}

	if (!V_ISARRAY(&pVarObj->variant)) {
	    warn(MY_VERSION ": Win32::OLE::Variant->Copy(): %d %s specified, "
		 "but variant is not a SAFEARRYA", items-1,
		 items > 2 ? "indices" : "index");
	    XSRETURN_EMPTY;
	}

	SAFEARRAY *psa = V_ISBYREF(pSource) ? *V_ARRAYREF(pSource)
	                                    : V_ARRAY(pSource);
	UINT cDims = SafeArrayGetDim(psa);
	if (items-1 != cDims) {
	    warn(MY_VERSION ": Win32::OLE::Variant->Copy() indices mismatch: "
		 "specified %d vs. required %d", items-1, cDims);
	    XSRETURN_EMPTY;
	}

	long *rgIndices;
	New(0, rgIndices, cDims, long);
	for (int iDim=0; iDim < cDims; ++iDim)
            rgIndices[iDim] = SvIV(ST(1+iDim));

	VARTYPE vt_base = V_VT(pSource) & ~VT_BYREF & ~VT_ARRAY;
	V_VT(&variant) = vt_base | VT_BYREF;
	V_VT(&byref) = vt_base;
	if (vt_base == VT_VARIANT)
            V_VARIANTREF(&variant) = &byref;
	else
            V_BYREF(&variant) = &V_BYREF(&byref);

	hr = SafeArrayGetElement(psa, rgIndices, V_BYREF(&variant));
	Safefree(rgIndices);
	if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	    XSRETURN_EMPTY;
	pSource = &variant;
    }

    WINOLEVARIANTOBJECT *pNewVar;
    Newz(0, pNewVar, 1, WINOLEVARIANTOBJECT);
    VariantInit(&pNewVar->variant);
    VariantInit(&pNewVar->byref);

    if (ix == 0)
	hr = VariantCopyInd(&pNewVar->variant, pSource);
    else
	hr = VariantCopy(&pNewVar->variant, pSource);

    VariantClear(&byref);
    if (FAILED(hr)) {
	Safefree(pNewVar);
	ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pNewVar,
		     WINOLEVARIANT_MAGIC);

    HV *stash = GetStash(PERL_OBJECT_THIS_ self);
    SV *sv = newSViv((IV)pNewVar);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
Dim(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);
    if (pVarObj == NULL)
	XSRETURN_EMPTY;

    VARIANT *pVariant = &pVarObj->variant;
    if (!V_ISARRAY(pVariant)) {
	warn(MY_VERSION ": Win32::OLE::Variant->Dim(): Variant type (0x%x) "
	     "is not an array", V_VT(pVariant));
	XSRETURN_EMPTY;
    }

    SAFEARRAY *psa;
    if (V_ISBYREF(pVariant))
	psa = *V_ARRAYREF(pVariant);
    else
	psa = V_ARRAY(pVariant);

    HRESULT hr = S_OK;
    UINT cDims = SafeArrayGetDim(psa);
    for (int iDim=0; iDim < cDims; ++iDim) {
	long lLBound, lUBound;
	hr = SafeArrayGetLBound(psa, 1+iDim, &lLBound);
	if (FAILED(hr))
	    break;
	hr = SafeArrayGetUBound(psa, 1+iDim, &lUBound);
	if (FAILED(hr))
	    break;
	AV *av = newAV();
	av_push(av, newSViv(lLBound));
	av_push(av, newSViv(lUBound));
	XPUSHs(sv_2mortal(newRV_noinc((SV*)av)));
    }

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    /* return list of array refs on stack */
}

void
Get(self,...)
    SV *self
ALIAS:
    Put = 1
PPCODE:
{
    char *paszMethod[] = {"Get", "Put"};
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);
    if (pVarObj == NULL)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    VARIANT *pVariant = &pVarObj->variant;

    if (!V_ISARRAY(pVariant)) {
	if (items-1 != ix) {
	    warn(MY_VERSION ": Win32::OLE::Variant->%s(): Wrong number of "
		 "arguments" , paszMethod[ix]);
	    XSRETURN_EMPTY;
	}
	if (ix == 0) { /* Get */
	    ST(0) = sv_newmortal();
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ pVariant, ST(0), olestash);
	    XSRETURN(1);
	}
	/* Put */
	AssignVariantFromSV(PERL_OBJECT_THIS_ ST(1), pVariant, olestash);
	XSRETURN_EMPTY;
    }

    SAFEARRAY *psa = V_ISBYREF(pVariant) ? *V_ARRAYREF(pVariant)
	                                  : V_ARRAY(pVariant);
    UINT cDims = SafeArrayGetDim(psa);

    /* Special case for one-dimensional VT_UI1 arrays */
    VARTYPE vt_base = V_VT(pVariant) & ~VT_BYREF & ~VT_ARRAY;
    if (vt_base == VT_UI1 && cDims == 1 && items-1 == ix) {
	if (ix == 0) { /* Get */
	    ST(0) = sv_newmortal();
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ &pVarObj->variant, ST(0),
			       olestash);
	    XSRETURN(1);
	}
	else { /* Put */
	    AssignVariantFromSV(PERL_OBJECT_THIS_ ST(1), pVariant, olestash);
	    XSRETURN_EMPTY;
	}
    }

    if (items-1 != cDims+ix) {
	warn(MY_VERSION ": Win32::OLE::Variant->%s(): Wrong number of indices; "
	     " dimension of SafeArray is %d", paszMethod[ix], cDims);
	XSRETURN_EMPTY;
    }

    ST(0) = &PL_sv_undef;
    long *rgIndices;
    New(0, rgIndices, cDims, long);
    for (int iDim=0; iDim < cDims; ++iDim)
        rgIndices[iDim] = SvIV(ST(1+iDim));

    VARIANT variant, byref;
    VariantInit(&variant);
    VariantInit(&byref);
    V_VT(&variant) = vt_base | VT_BYREF;
    V_VT(&byref) = vt_base;
    if (vt_base == VT_VARIANT)
        V_VARIANTREF(&variant) = &byref;
    else
        V_BYREF(&variant) = &V_BYREF(&byref);

    HRESULT hr = S_OK;
    if (ix == 0) { /* Get */
	hr = SafeArrayGetElement(psa, rgIndices, V_BYREF(&variant));
	if (SUCCEEDED(hr)) {
	    ST(0) = sv_newmortal();
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ &variant, ST(0), olestash);
	}
    }
    else { /* Put */
	AssignVariantFromSV(PERL_OBJECT_THIS_ ST(items-1), &variant, olestash);
	hr = SafeArrayPutElement(psa, rgIndices, V_BYREF(&variant));
    }
    VariantClear(&byref);
    Safefree(rgIndices);
    CheckOleError(PERL_OBJECT_THIS_ olestash, hr);
    XSRETURN(1);
}

void
LastError(self,...)
    SV *self
PPCODE:
{
    // Win32::OLE::Variant->LastError() exists only for backward compatibility.
    // It is now just a proxy for Win32::OLE->LastError().

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    SV *sv = items == 1 ? NULL : ST(1);

    PUSHMARK(sp);
    PUSHs(sv_2mortal(newSVpv(HvNAME(olestash), 0)));
    if (sv)
	PUSHs(sv);
    PUTBACK;
    perl_call_method("LastError", GIMME_V);
    SPAGAIN;

    // return whatever Win32::OLE->LastError() returned
}

void
Type(self)
    SV *self
ALIAS:
    Value = 1
    _Value = 2
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj != NULL) {
	HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
	SetLastOleError(PERL_OBJECT_THIS_ olestash);
	ST(0) = sv_newmortal();
	if (ix == 0) /* Type */
	    sv_setiv(ST(0), V_VT(&pVarObj->variant));
	else if (ix == 1) /* Value */
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ &pVarObj->variant, ST(0),
			       olestash);
	else if (ix == 2) /* _Value, see also: _Clone (alias of Copy) */
	    SetSVFromVariantEx(PERL_OBJECT_THIS_ &pVarObj->variant, ST(0),
			       olestash, TRUE);
    }
    XSRETURN(1);
}

void
Unicode(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj != NULL) {
	VARIANT Variant;
	VARIANT *pVariant = &pVarObj->variant;
	HRESULT hr = S_OK;

	HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
	SetLastOleError(PERL_OBJECT_THIS_ olestash);
	VariantInit(&Variant);
	if ((V_VT(pVariant) & ~VT_BYREF) != VT_BSTR) {
	    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ olestash,
				    LCID_NAME, LCID_LEN, lcidDefault);

	    hr = VariantChangeTypeEx(&Variant, pVariant, lcid, 0, VT_BSTR);
	    pVariant = &Variant;
	}

	if (!CheckOleError(PERL_OBJECT_THIS_ olestash, hr)) {
	    BSTR bstr = V_ISBYREF(pVariant) ? *V_BSTRREF(pVariant)
		                            : V_BSTR(pVariant);
	    STRLEN olecharlen = SysStringLen(bstr);
	    SV *sv = newSVpv((char*)bstr, 2*olecharlen);
	    STRLEN len;
	    U16 *pus = (U16 *)SvPV(sv, len);
	    for (STRLEN i=0 ; i < olecharlen ; ++i)
		pus[i] = htons(pus[i]);

	    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv),
					gv_stashpv("Unicode::String", TRUE)));
	}
	VariantClear(&Variant);
    }
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::NLS

void
CompareString(lcid,flags,str1,str2)
    IV lcid
    IV flags
    SV *str1
    SV *str2
PPCODE:
{
    STRLEN length1;
    STRLEN length2;
    char *string1 = SvPV(str1, length1);
    char *string2 = SvPV(str2, length2);

    int res = CompareStringA(lcid, flags, string1, length1, string2, length2);
    XSRETURN_IV(res);
}

void
LCMapString(lcid,flags,str)
    IV lcid
    IV flags
    SV *str
PPCODE:
{
    SV *sv = sv_newmortal();
    STRLEN length;
    char *string = SvPV(str,length);
    int len = LCMapStringA(lcid, flags, string, length, NULL, 0);
    if (len > 0) {
	SvUPGRADE(sv, SVt_PV);
	SvGROW(sv, len+1);
	SvCUR(sv) = LCMapStringA(lcid, flags, string, length,
				 SvPVX(sv), SvLEN(sv));
	if (SvCUR(sv))
	    SvPOK_on(sv);
    }
    ST(0) = sv;
    XSRETURN(1);
}

void
GetLocaleInfo(lcid,lctype)
    IV lcid
    IV lctype
PPCODE:
{
    SV *sv = sv_newmortal();
    int len = GetLocaleInfoA(lcid, lctype, NULL, 0);
    if (len > 0) {
	SvUPGRADE(sv, SVt_PV);
	SvGROW(sv, len);
	SvCUR(sv) = GetLocaleInfoA(lcid, lctype, SvPVX(sv), SvLEN(sv));
	if (SvCUR(sv)) {
	    -- SvCUR(sv);
	    SvPOK_on(sv);
	}
    }
    ST(0) = sv;
    XSRETURN(1);
}

void
GetStringType(lcid,type,str)
    IV lcid
    IV type
    SV *str
PPCODE:
{
    STRLEN len;
    char *string = SvPV(str, len);
    unsigned short *pCharType;

    New(0, pCharType, len, unsigned short);
    if (GetStringTypeA(lcid, type, string, len, pCharType)) {
	EXTEND(sp, len);
	for (int i=0 ; i < len ; ++i)
	    PUSHs(sv_2mortal(newSViv(pCharType[i])));
    }
    Safefree(pCharType);
}

void
GetSystemDefaultLangID()
PPCODE:
{
    LANGID langID = GetSystemDefaultLangID();
    if (langID != 0) {
	EXTEND(sp, 1);
	XSRETURN_IV(langID);
    }
}

void
GetSystemDefaultLCID()
PPCODE:
{
    LCID lcid = GetSystemDefaultLCID();
    if (lcid != 0) {
	EXTEND(sp, 1);
	XSRETURN_IV(lcid);
    }
}

void
GetUserDefaultLangID()
PPCODE:
{
    LANGID langID = GetUserDefaultLangID();
    if (langID != 0) {
	EXTEND(sp, 1);
	XSRETURN_IV(langID);
    }
}

void
GetUserDefaultLCID()
PPCODE:
{
    LCID lcid = GetUserDefaultLCID();
    if (lcid != 0) {
	EXTEND(sp, 1);
	XSRETURN_IV(lcid);
    }
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::TypeLib

void
_new(self,object)
    SV *self
    SV *object
PPCODE:
{
    ITypeLib  *pTypeLib;
    ITypeInfo *pTypeInfo;
    TLIBATTR  *pTLibAttr;

    // XXX object should be a typelib, not a Win32::OLE object!

    WINOLEOBJECT *pOleObj = GetOleObject(PERL_OBJECT_THIS_ object);
    if (pOleObj == NULL)
        XSRETURN_EMPTY;

    unsigned int count;
    HRESULT hr = pOleObj->pDispatch->GetTypeInfoCount(&count);
    HV *stash = SvSTASH(pOleObj->self);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) || count == 0)
        XSRETURN_EMPTY;

    hr = pOleObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    unsigned int index;
    hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
    pTypeInfo->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
        XSRETURN_EMPTY;

    hr = pTypeLib->GetLibAttr(&pTLibAttr);
    if (FAILED(hr)) {
	pTypeLib->Release();
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeLibObject(PERL_OBJECT_THIS_ pTypeLib,
					   pTLibAttr));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(PERL_OBJECT_THIS_ self);
    if (pObj != NULL) {
	RemoveFromObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pObj);
	if (pObj->pTypeLib != NULL) {
	    pObj->pTypeLib->ReleaseTLibAttr(pObj->pTLibAttr);
	    pObj->pTypeLib->Release();
	}
	Safefree(pObj);
    }
    XSRETURN_EMPTY;
}

void
_GetDocumentation(self,index=-1)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    DWORD dwHelpContext;
    BSTR bstrName, bstrDocString, bstrHelpFile;
    HRESULT hr = pObj->pTypeLib->GetDocumentation(index, &bstrName,
			  &bstrDocString, &dwHelpContext, &bstrHelpFile);
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = GetDocumentation(PERL_OBJECT_THIS_ bstrName, bstrDocString,
			      dwHelpContext, bstrHelpFile);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetLibAttr(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    TLIBATTR *p = pObj->pTLibAttr;
    HV *hv = newHV();

    hv_store(hv, "lcid",          4, newSViv(p->lcid), 0);
    hv_store(hv, "syskind",       7, newSViv(p->syskind), 0);
    hv_store(hv, "wLibFlags",     9, newSViv(p->wLibFlags), 0);
    hv_store(hv, "wMajorVerNum", 12, newSViv(p->wMajorVerNum), 0);
    hv_store(hv, "wMinorVerNum", 12, newSViv(p->wMinorVerNum), 0);
    hv_store(hv, "guid",          4, SetSVFromGUID(PERL_OBJECT_THIS_
						   p->guid), 0);

    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetTypeInfoCount(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    XSRETURN_IV(pObj->pTypeLib->GetTypeInfoCount());
}

void
_GetTypeInfo(self,index)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeLib->GetTypeInfo(index, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(PERL_OBJECT_THIS_ pTypeInfo,
					    pTypeAttr));
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::TypeInfo

void
_new(self,object)
    SV *self
    SV *object
PPCODE:
{
    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    WINOLEOBJECT *pOleObj = GetOleObject(PERL_OBJECT_THIS_ object);
    if (pOleObj == NULL)
        XSRETURN_EMPTY;

    unsigned int count;
    HRESULT hr = pOleObj->pDispatch->GetTypeInfoCount(&count);
    HV *olestash = SvSTASH(pOleObj->self);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr) || count == 0)
        XSRETURN_EMPTY;

    hr = pOleObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
        XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(PERL_OBJECT_THIS_ pTypeInfo,
					    pTypeAttr));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj != NULL) {
	RemoveFromObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pObj);
	if (pObj->pTypeInfo != NULL) {
	    pObj->pTypeInfo->ReleaseTypeAttr(pObj->pTypeAttr);
	    pObj->pTypeInfo->Release();
	}
	Safefree(pObj);
    }
    XSRETURN_EMPTY;
}

void
_GetContainingTypeLib(self)
    SV *self
PPCODE:
{
    ITypeLib  *pTypeLib;
    TLIBATTR  *pTLibAttr;

    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    unsigned int index;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
        XSRETURN_EMPTY;

    hr = pTypeLib->GetLibAttr(&pTLibAttr);
    if (FAILED(hr)) {
	pTypeLib->Release();
	ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeLibObject(PERL_OBJECT_THIS_ pTypeLib,
					   pTLibAttr));
    XSRETURN(1);
}

void
_GetDocumentation(self,memid=-1)
    SV *self
    IV memid
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    DWORD dwHelpContext;
    BSTR bstrName, bstrDocString, bstrHelpFile;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetDocumentation(memid, &bstrName,
			   &bstrDocString, &dwHelpContext, &bstrHelpFile);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = GetDocumentation(PERL_OBJECT_THIS_ bstrName, bstrDocString,
			      dwHelpContext, bstrHelpFile);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetFuncDesc(self,index)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    FUNCDESC *p;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetFuncDesc(index, &p);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = newHV();
    hv_store(hv, "memid",         5, newSViv(p->memid), 0);
    // /* [size_is] */ SCODE __RPC_FAR *lprgscode;
    hv_store(hv, "funckind",      8, newSViv(p->funckind), 0);
    hv_store(hv, "invkind",       7, newSViv(p->invkind), 0);
    hv_store(hv, "callconv",      8, newSViv(p->callconv), 0);
    hv_store(hv, "cParams",       7, newSViv(p->cParams), 0);
    hv_store(hv, "cParamsOpt",   10, newSViv(p->cParamsOpt), 0);
    hv_store(hv, "oVft",          4, newSViv(p->oVft), 0);
    hv_store(hv, "cScodes",       7, newSViv(p->cScodes), 0);
    hv_store(hv, "wFuncFlags",   10, newSViv(p->wFuncFlags), 0);

    HV *elemdesc = TranslateElemDesc(PERL_OBJECT_THIS_ &p->elemdescFunc,
				     pObj, olestash);
    hv_store(hv, "elemdescFunc", 12, newRV_noinc((SV*)elemdesc), 0);

    if (p->cParams > 0) {
	AV *av = newAV();

	for (int i = 0; i < p->cParams; ++i) {
	    elemdesc = TranslateElemDesc(PERL_OBJECT_THIS_
					 &p->lprgelemdescParam[i],
					 pObj, olestash);
	    av_push(av, newRV_noinc((SV*)elemdesc));
	}
	hv_store(hv, "rgelemdescParam", 15, newRV_noinc((SV*)av), 0);
    }

    pObj->pTypeInfo->ReleaseFuncDesc(p);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetImplTypeFlags(self,index)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    int flags;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetImplTypeFlags(index, &flags);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    XSRETURN_IV(flags);
}

void
_GetImplTypeInfo(self,index)
    SV *self
    IV index
PPCODE:
{
    HREFTYPE  hRefType;
    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetRefTypeOfImplType(index, &hRefType);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pObj->pTypeInfo->GetRefTypeInfo(hRefType, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(PERL_OBJECT_THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    New(0, pObj, 1, WINOLETYPEINFOOBJECT);
    pObj->pTypeInfo = pTypeInfo;
    pObj->pTypeAttr = pTypeAttr;

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pObj,
		     WINOLETYPEINFO_MAGIC);

    SV *sv = newSViv((IV)pObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv),
				GetStash(PERL_OBJECT_THIS_ self)));
    XSRETURN(1);
}

void
_GetNames(self,memid,count)
    SV *self
    IV memid
    IV count
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    BSTR *rgbstr;
    New(0, rgbstr, count, BSTR);
    unsigned int cNames;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetNames(memid, rgbstr, count, &cNames);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    AV *av = newAV();
    for (int i = 0 ; i < cNames ; ++i) {
	char szName[32];
	// XXX use correct codepage ???
	char *pszName = GetMultiByte(PERL_OBJECT_THIS_ rgbstr[i],
				     szName, sizeof(szName), CP_ACP);
	SysFreeString(rgbstr[i]);
	av_push(av, newSVpv(pszName, 0));
	ReleaseBuffer(PERL_OBJECT_THIS_ pszName, szName);
    }
    Safefree(rgbstr);

    ST(0) = sv_2mortal(newRV_noinc((SV*)av));
    XSRETURN(1);
}

void
_GetTypeAttr(self)
    SV *self
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    TYPEATTR *p = pObj->pTypeAttr;
    HV *hv = newHV();

    hv_store(hv, "guid",              4, SetSVFromGUID(PERL_OBJECT_THIS_
						       p->guid), 0);
    hv_store(hv, "lcid",              4, newSViv(p->lcid), 0);
    hv_store(hv, "memidConstructor", 16, newSViv(p->memidConstructor), 0);
    hv_store(hv, "memidDestructor",  15, newSViv(p->memidDestructor), 0);
    hv_store(hv, "typekind",          8, newSViv(p->typekind), 0);
    hv_store(hv, "cFuncs",            6, newSViv(p->cFuncs), 0);
    hv_store(hv, "cVars",             5, newSViv(p->cVars), 0);
    hv_store(hv, "cImplTypes",       10, newSViv(p->cImplTypes), 0);
    hv_store(hv, "cbSizeVft",         9, newSViv(p->cbSizeVft), 0);
    hv_store(hv, "wTypeFlags",       10, newSViv(p->wTypeFlags), 0);
    hv_store(hv, "wMajorVerNum",     12, newSViv(p->wMajorVerNum), 0);
    hv_store(hv, "wMinorVerNum",     12, newSViv(p->wMinorVerNum), 0);
    //TYPEDESC tdescAlias;	  // If TypeKind == TKIND_ALIAS,
    //                            // specifies the type for which
    //                            // this type is an alias.
    //IDLDESC idldescType;	  // IDL attributes of the
    //                            // described type.


    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetVarDesc(self,index)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(PERL_OBJECT_THIS_ self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    VARDESC *p;
    HV *olestash = GetWin32OleStash(PERL_OBJECT_THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetVarDesc(index, &p);
    if (CheckOleError(PERL_OBJECT_THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = newHV();
    hv_store(hv, "memid",        5, newSViv(p->memid), 0);
    // LPOLESTR lpstrSchema;
    hv_store(hv, "wVarFlags",    9, newSViv(p->wVarFlags), 0);
    hv_store(hv, "varkind",      7, newSViv(p->varkind), 0);

    HV *elemdesc = TranslateElemDesc(PERL_OBJECT_THIS_ &p->elemdescVar,
				     pObj, olestash);
    hv_store(hv, "elemdescVar", 11, newRV_noinc((SV*)elemdesc), 0);

    if (p->varkind == VAR_PERINSTANCE)
	hv_store(hv, "oInst",    5, newSViv(p->oInst), 0);

    if (p->varkind == VAR_CONST) {
	// XXX should be stored as a Win32::OLE::Variant object ?
	SV *sv = newSVpv("",0);
	SetSVFromVariantEx(PERL_OBJECT_THIS_ p->lpvarValue, sv, olestash);
	hv_store(hv, "varValue", 8, sv, 0);
    }

    pObj->pTypeInfo->ReleaseVarDesc(p);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

