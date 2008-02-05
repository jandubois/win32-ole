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
 * File structure:
 *
 * - C helper routines
 * - Package Win32::OLE             Constructor and method invocation
 * - Package Win32::OLE::Tie        Implements properties as tied hash
 * - Package Win32::OLE::Const      Load application constants from type library
 * - Package Win32::OLE::Enum       OLE collection enumeration
 * - Package Win32::OLE::Variant    Implements Perl VARIANT objects
 * - Package Win32::OLE::NLS        National Language Support
 *
 */

#ifdef XS_VERSION
#   define MY_VERSION "Win32::OLE(" XS_VERSION ")"
#else
#   define MY_VERSION "Win32::OLE(?.??)"
#endif

#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#define _WIN32_DCOM
#include <windows.h>

#ifdef _DEBUG
#   include <crtdbg.h>
#   define DEBUGBREAK _CrtDbgBreak()
#else
#   define DEBUGBREAK
#endif

#if defined(__cplusplus)
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

#if (PATCHLEVEL < 5) && !defined(PL_dowarn)
#   define PL_hints	hints
#   define PL_dowarn	dowarn
#   define PL_modglobal	modglobal
#   define PL_sv_undef	sv_undef
#   define PL_sv_yes    sv_yes
#   define PL_sv_no     sv_no
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
    HV           *stash;

}   WINOLEENUMOBJECT;

/* Win32::OLE::Variant object */
typedef struct
{
    OBJECTHEADER header;

    VARIANT variant;
    VARIANT byref;

    HV      *stash;    /* for VT_DISPATCH and VT_UNKNOWN Variants */

}   WINOLEVARIANTOBJECT;

/* forward declarations */
HRESULT SetSVFromVariant(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash);

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

IV
QueryPkgVar(CPERLarg_ HV *stash, char *var, STRLEN len, IV def)
{
    SV *sv;
    GV **gv = (GV **) hv_fetch(stash, var, len, FALSE);

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

    IV OleWarn = QueryPkgVar(PERL_OBJECT_THIS_ stash, WARN_NAME, WARN_LEN, 0);
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
    
    if (OleWarn > 1 || (OleWarn == 1 && PL_dowarn)) {
	PUSHMARK(sp) ;
	XPUSHs(sv);
	PUTBACK;
	perl_call_pv(OleWarn < 3 ? "Carp::carp" : "Carp::croak", G_DISCARD);
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

    warn("%s: DESTROY must be a method name or a CODE reference", szMethod);
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
    if (destroy !=NULL) {
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
	hr = SetSVFromVariant(PERL_OBJECT_THIS_ &variant, sv, stash);
    }
    VariantClear(&variant);
    if (FAILED(hr)) {
        SvREFCNT_dec(sv);
	sv = &PL_sv_undef;
	ReportOleError(PERL_OBJECT_THIS_ stash, hr);
    }
    return sv;

}   /* NextEnumElement */

HRESULT
SetVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp)
{
    HRESULT hr = S_OK;
    VariantInit(pVariant);

    /* XXX requirement to call mg_get() may change in Perl > 5.004 */
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

#define SETiVRETURN(x,f)					\
		    if (x->vt&VT_BYREF) {			\
			sv_setiv(sv, (long)*V_##f##REF(x));	\
		    } else {					\
			sv_setiv(sv, (long)V_##f(x));		\
		    }

#define SETnVRETURN(x,f)					\
		    if (x->vt&VT_BYREF) {			\
			sv_setnv(sv, (double)*V_##f##REF(x));	\
		    } else {					\
			sv_setnv(sv, (double)V_##f(x));		\
		    }

HRESULT
SetSVFromVariant(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash)
{
    HRESULT hr = S_OK;
    sv_setsv(sv, &PL_sv_undef);

    if (V_ISARRAY(pVariant)) {
	SAFEARRAY *psa = V_ISBYREF(pVariant) ? *V_ARRAYREF(pVariant)
	                                     : V_ARRAY(pVariant);
	AV **pav;
	IV index;
	long *pArrayIndex, *pLowerBound, *pUpperBound;
	VARIANT variant;

	int dim = SafeArrayGetDim(psa);

	VariantInit(&variant);
	V_VT(&variant) = (V_VT(pVariant) & ~VT_ARRAY) | VT_BYREF;

	/* convert 1-dim UI1 ARRAY to simple SvPV */
	if (dim == 1 && (V_VT(pVariant) & ~VT_ARRAY & ~VT_BYREF) == VT_UI1) {
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
		hr = SetSVFromVariant(PERL_OBJECT_THIS_ &variant, val, stash);
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

    while (V_VT(pVariant) == (VT_VARIANT|VT_BYREF))
	pVariant = V_VARIANTREF(pVariant);

    switch(V_VT(pVariant) & ~VT_BYREF)
    {
    case VT_VARIANT: /* invalid, should never happen */
    case VT_EMPTY:
    case VT_NULL:
	/* return "undef" */
	break;

    case VT_UI1:
	SETiVRETURN(pVariant, UI1);
	break;

    case VT_I2:
	SETiVRETURN(pVariant, I2);
	break;

    case VT_I4:
	SETiVRETURN(pVariant, I4);
	break;

    case VT_R4:
	SETnVRETURN(pVariant, R4);
	break;

    case VT_R8:
	SETnVRETURN(pVariant, R8);
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
	SETiVRETURN(pVariant, ERROR);
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

}   /* SetSVFromVariant */

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
Initialize(CPERLarg_ DWORD dwCoInit=COINIT_MULTITHREADED)
{
    dPERINTERP;

    DBG(("Initialize\n"));
    EnterCriticalSection(&g_CriticalSection);

    if (!g_bInitialized)
    {
	DBG(("(Co|Ole)Initialize(Ex)?\n"));
	if (dwCoInit == -1) {
	    OleInitialize(NULL);
	    g_pfnCoUninitialize = &OleUninitialize;
	}
	else {
	    if (g_pfnCoInitializeEx == NULL)
		CoInitialize(NULL);
	    else
		g_pfnCoInitializeEx(NULL, dwCoInit);
	    g_pfnCoUninitialize = &CoUninitialize;
	}

	g_bInitialized = TRUE;
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

	    default:
		DBG(("Unknown magic number: %08lx", g_pObj->lMagic));
		break;
	    }
	    g_pObj = g_pObj->pNext;
	}

	SpinMessageLoop();
	DBG(("CoUninitialize\n"));
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

#if defined(__cplusplus)
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

    switch (ix)
    {
    case 0:
    {
	DWORD dwCoInit = COINIT_MULTITHREADED;

	if (items > 1 && SvOK(ST(1)))
	    dwCoInit = SvIV(ST(1));

	Initialize(PERL_OBJECT_THIS_ dwCoInit);
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

    XSRETURN_UNDEF;
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
	XSRETURN_UNDEF;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(PERL_OBJECT_THIS);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(PERL_OBJECT_THIS_ ST(2),
				       "Win32::OLE::new");

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
		XSRETURN_UNDEF;
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
	    XSRETURN_UNDEF;
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
		pVarObj->stash = stash; /* XXX refcount ??? */
		ST(0) = &PL_sv_yes;
	    }
	}
	else {
	    hr = SetSVFromVariant(PERL_OBJECT_THIS_ &result, retval, stash);
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

    if (items != 2) {
	warn("Usage: Win32::OLE->GetActiveObject(progid)");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    if (!SvPOK(self)) {
	warn("Win32::OLE->GetActiveObject: Must be called as a class method");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    Initialize(PERL_OBJECT_THIS);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    buffer = SvPV(progid, length);
    pBuffer = GetWideChar(PERL_OBJECT_THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    if (isalpha(pBuffer[0]))
        hr = CLSIDFromProgID(pBuffer, &clsid);
    else
        hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

    hr = GetActiveObject(clsid, 0, &pUnknown);
    /* Don't call CheckOleError! Return "undef" for "Server not running" */
    if (FAILED(hr))
	XSRETURN_UNDEF;

    hr = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, NULL);
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

    if (items != 2) {
	warn("Usage: Win32::OLE->GetObject(pathname)");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *pathname = ST(1);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    if (!SvPOK(self)) {
	warn("Win32::OLE->GetObject: Must be called as a class method");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    Initialize(PERL_OBJECT_THIS);
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    hr = CreateBindCtx(0, &pBindCtx);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

    buffer = SvPV(pathname, len);
    pBuffer = GetWideChar(PERL_OBJECT_THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    hr = MkParseDisplayName(pBindCtx, pBuffer, &ulEaten, &pMoniker);
    ReleaseBuffer(PERL_OBJECT_THIS_ pBuffer, Buffer);
    if (FAILED(hr)) {
	pBindCtx->Release();
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "after character %lu in \"%s\"", ulEaten, buffer);
	ReportOleError(PERL_OBJECT_THIS_ stash, hr, NULL, sv);
	XSRETURN_UNDEF;
    }

    hr = pMoniker->BindToObject(pBindCtx, NULL, IID_IDispatch,
				 (void**)&pDispatch);
    pBindCtx->Release();
    pMoniker->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(PERL_OBJECT_THIS_ stash, pDispatch, NULL);
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
	XSRETURN_UNDEF;
    }

    SV *object = ST(1);

    if (!sv_isobject(object) || !sv_derived_from(object, szWINOLE))
	XSRETURN_UNDEF;

    WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ object);
    if (pObj == NULL)
	XSRETURN_UNDEF;

    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int count;
    BSTR bstr;

    HRESULT hr = pObj->pDispatch->GetTypeInfoCount(&count);
    if (FAILED(hr) || count == 0)
	XSRETURN_UNDEF;

    HV *stash = gv_stashsv(ST(0), TRUE);
    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			    lcidDefault);
    UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    SetLastOleError(PERL_OBJECT_THIS_ stash);
    hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

    /* Return ('TypeLib Name', 'Class Name') in array context */
    if (GIMME_V == G_ARRAY) {
	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	    XSRETURN_UNDEF;
	}

	hr = pTypeLib->GetDocumentation(-1, &bstr, NULL, NULL, NULL);
	pTypeLib->Release();
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
	    XSRETURN_UNDEF;
	}

	PUSHs(sv_2mortal(sv_setwide(PERL_OBJECT_THIS_ NULL, bstr, cp)));
	SysFreeString(bstr);
    }

    hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstr, NULL, NULL, NULL);
    pTypeInfo->Release();
    if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	XSRETURN_UNDEF;

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
	hr = SetSVFromVariant(PERL_OBJECT_THIS_ &result, ST(0), stash);
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
	XSRETURN_UNDEF;

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

    Initialize(PERL_OBJECT_THIS);
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
	    XSRETURN_UNDEF;

	stash = SvSTASH(pObj->self);
	hr = pObj->pDispatch->GetTypeInfoCount(&count);
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr) || count == 0)
	    XSRETURN_UNDEF;

	lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
			   lcidDefault);
	cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN, cpDefault);

	hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_UNDEF;

	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	pTypeInfo->Release();
	if (CheckOleError(PERL_OBJECT_THIS_ stash, hr))
	    XSRETURN_UNDEF;
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
	    XSRETURN_UNDEF;

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
	    XSRETURN_UNDEF;
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
		hr = SetSVFromVariant(PERL_OBJECT_THIS_ pVarDesc->lpvarValue,
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
	    pEnumObj->stash = SvSTASH(pObj->self);
	    SetLastOleError(PERL_OBJECT_THIS_ pEnumObj->stash);
	    pEnumObj->pEnum = CreateEnumVARIANT(PERL_OBJECT_THIS_ pObj);
	}
    }
    else { /* Clone */
	WINOLEENUMOBJECT *pOriginal = GetOleEnumObject(PERL_OBJECT_THIS_ self);
	if (pOriginal != NULL) {
	    HRESULT hr = pOriginal->pEnum->Clone(&pEnumObj->pEnum);
	    SetLastOleError(PERL_OBJECT_THIS_ pOriginal->stash);
	    CheckOleError(PERL_OBJECT_THIS_ pOriginal->stash, hr);
	    pEnumObj->stash = pOriginal->stash;
	}
    }

    if (pEnumObj->pEnum == NULL) {
	Safefree(pEnumObj);
	XSRETURN_UNDEF;
    }

    SvREFCNT_inc(pEnumObj->stash);
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
	SvREFCNT_dec(pEnumObj->stash);
	if (pEnumObj->pEnum != NULL)
	    pEnumObj->pEnum->Release();
	Safefree(pEnumObj);
    }
    XSRETURN_EMPTY;
}

void
Next(self,...)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(PERL_OBJECT_THIS_ self);
    if (pEnumObj == NULL)
	XSRETURN_UNDEF;

    int count = (items > 1) ? SvIV(ST(1)) : 1;
    if (count < 1) {
	warn(MY_VERSION ": Win32::OLE::Enum::Next: invalid Count %ld", count);
	DEBUGBREAK;
	count = 1;
    }

    SetLastOleError(PERL_OBJECT_THIS_ pEnumObj->stash);

    SV *sv = NULL;
    while (count-- > 0) {
	sv = NextEnumElement(PERL_OBJECT_THIS_ pEnumObj->pEnum,
			     pEnumObj->stash);
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

    SetLastOleError(PERL_OBJECT_THIS_ pEnumObj->stash);
    HRESULT hr = pEnumObj->pEnum->Reset();
    CheckOleError(PERL_OBJECT_THIS_ pEnumObj->stash, hr);
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

    SetLastOleError(PERL_OBJECT_THIS_ pEnumObj->stash);
    int count = (items > 1) ? SvIV(ST(1)) : 1;
    HRESULT hr = pEnumObj->pEnum->Skip(count);
    CheckOleError(PERL_OBJECT_THIS_ pEnumObj->stash, hr);
    ST(0) = boolSV(hr == S_OK);
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Variant

void
new(self,vt,data)
    SV *self
    IV vt
    SV *data
PPCODE:
{
    char *ptr;
    STRLEN length;
    HV *stash = GetStash(PERL_OBJECT_THIS_ self);
    HRESULT hr;
    WINOLEVARIANTOBJECT *pVarObj;

    // XXX Initialize should be superfluous here
    // Initialize();
    SetLastOleError(PERL_OBJECT_THIS_ stash);

    New(0, pVarObj, 1, WINOLEVARIANTOBJECT);
    VariantInit(&pVarObj->variant);
    VariantInit(&pVarObj->byref);
    pVarObj->stash = stash; /* XXX refcnt ??? */
    V_VT(&pVarObj->variant) = vt & ~VT_BYREF;

    /* XXX requirement to call mg_get() may change in Perl > 5.004 */
    if (SvGMAGICAL(data))
	mg_get(data);

    switch (V_VT(&pVarObj->variant)) {
    case VT_EMPTY:
    case VT_NULL:
	break;

    case VT_I2:
	V_I2(&pVarObj->variant) = SvIV(data);
	break;

    case VT_I4:
	V_I4(&pVarObj->variant) = SvIV(data);
	break;

    case VT_R4:
	V_R4(&pVarObj->variant) = SvNV(data);
	break;

    case VT_R8:
	V_R8(&pVarObj->variant) = SvNV(data);
	break;

    case VT_CY:
    case VT_DATE:
    {
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
				lcidDefault);
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);

	V_VT(&pVarObj->variant) = VT_BSTR;
	ptr = SvPV(data, length);
	V_BSTR(&pVarObj->variant) = AllocOleString(PERL_OBJECT_THIS_ ptr,
						   length, cp);
	VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant, lcid,0, vt);
	break;
    }

    case VT_BSTR:
    {
	UINT cp = QueryPkgVar(PERL_OBJECT_THIS_ stash, CP_NAME, CP_LEN,
			      cpDefault);

	ptr = SvPV(data, length);
	V_BSTR(&pVarObj->variant) = AllocOleString(PERL_OBJECT_THIS_ ptr,
						   length, cp);
	break;
    }

    case VT_DISPATCH:
    {
	/* Argument MUST be a valid Perl OLE object! */
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ data);
	if (pObj == NULL)
	    V_VT(&pVarObj->variant) = VT_EMPTY;
	else {
	    pObj->pDispatch->AddRef();
	    V_DISPATCH(&pVarObj->variant) = pObj->pDispatch;
	    pVarObj->stash = SvSTASH(pObj->self); /* XXX refcnt ??? */
	}
	break;
    }

    case VT_ERROR:
	V_ERROR(&pVarObj->variant) = SvIV(data);
	break;

    case VT_BOOL:
	/* Either all bits are 0 or ALL bits MUST BE 1 */
	V_BOOL(&pVarObj->variant) = SvTRUE(data) ? ~0 : 0;
	break;

    /* case VT_VARIANT: invalid without VT_BYREF */

    case VT_UNKNOWN:
    {
	/* Argument MUST be a valid Perl OLE object! */
	/* Query IUnknown interface to allow identity tests */
	WINOLEOBJECT *pObj = GetOleObject(PERL_OBJECT_THIS_ data);
	if (pObj == NULL)
	    V_VT(&pVarObj->variant) = VT_EMPTY;
	else {
	    hr = pObj->pDispatch->QueryInterface(IID_IUnknown,
			           (void**)&V_UNKNOWN(&pVarObj->variant));
	    pVarObj->stash = SvSTASH(pObj->self); /* XXX refcnt ??? */
	    CheckOleError(PERL_OBJECT_THIS_ SvSTASH(pObj->self), hr);
	}
	break;
    }

    case VT_UI1:
	if (SvPOK(data)) {
	    unsigned char* pDest;

	    ptr = SvPV(data, length);
	    V_ARRAY(&pVarObj->variant) = SafeArrayCreateVector(VT_UI1, 0,
							       length);
	    if (V_ARRAY(&pVarObj->variant) != NULL) {
		V_VT(&pVarObj->variant) = VT_UI1 | VT_ARRAY;
		hr = SafeArrayAccessData(V_ARRAY(&pVarObj->variant),
						   (void**)&pDest);
		if (FAILED(hr)) {
		    VariantClear(&pVarObj->variant);
		    ReportOleError(PERL_OBJECT_THIS_ stash, hr);
		}
		else {
		    memcpy(pDest, ptr, length);
		    SafeArrayUnaccessData(V_ARRAY(&pVarObj->variant));
		}
	    }
	}
	else
	    V_UI1(&pVarObj->variant) = SvIV(data);

	break;

    default:
	warn(MY_VERSION ": Win32::OLE::Variant::new: Invalid value type %d",
	     V_VT(&pVarObj->variant));
	DEBUGBREAK;
	Safefree(pVarObj);
	XSRETURN_UNDEF;
    }

    if (vt & VT_BYREF) {
	pVarObj->byref = pVarObj->variant;
	VariantInit(&pVarObj->variant);
	V_VT(&pVarObj->variant) = vt;
	V_BYREF(&pVarObj->variant) = &V_UI1(&pVarObj->byref);
    }

    AddToObjectChain(PERL_OBJECT_THIS_ (OBJECTHEADER*)pVarObj,
		     WINOLEVARIANT_MAGIC);

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
	/* XXX dec refcnt(stash) ??? */
	Safefree(pVarObj);
    }

    XSRETURN_EMPTY;
}

void
Type(self)
    SV *self
ALIAS:
    Value = 1
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(PERL_OBJECT_THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj != NULL) {
	HV *stash = GetStash(PERL_OBJECT_THIS_ self);
	SetLastOleError(PERL_OBJECT_THIS_ stash); /* XXX */
	ST(0) = sv_newmortal();
	if (ix == 0)
	    sv_setiv(ST(0), V_VT(&pVarObj->variant));
	else
	    SetSVFromVariant(PERL_OBJECT_THIS_ &pVarObj->variant,
			     ST(0), pVarObj->stash);
    }
    XSRETURN(1);
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
	HV *stash = GetStash(PERL_OBJECT_THIS_ self);
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
				lcidDefault);

	SetLastOleError(PERL_OBJECT_THIS_ stash);
	VariantInit(&variant);
	hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, type);
	if (SUCCEEDED(hr)) {
	    ST(0) = sv_newmortal();
	    SetSVFromVariant(PERL_OBJECT_THIS_ &variant, ST(0), pVarObj->stash);
	}
	VariantClear(&variant);
	CheckOleError(PERL_OBJECT_THIS_ stash, hr);
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
	HV *stash = GetStash(PERL_OBJECT_THIS_ self);
	LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash, LCID_NAME, LCID_LEN,
				lcidDefault);

	SetLastOleError(PERL_OBJECT_THIS_ stash);
	/* XXX: Does it work with VT_BYREF? */
	hr = VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant,
				  lcid, 0, type);
	CheckOleError(PERL_OBJECT_THIS_ stash, hr);
    }

    if (FAILED(hr))
	ST(0) = &PL_sv_undef;

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
	HV *stash = GetStash(PERL_OBJECT_THIS_ self);
	VARIANT Variant;
	VARIANT *pVariant = &pVarObj->variant;
	HRESULT hr = S_OK;

	SetLastOleError(PERL_OBJECT_THIS_ stash);
	VariantInit(&Variant);
	if ((V_VT(pVariant) & ~VT_BYREF) != VT_BSTR) {
	    LCID lcid = QueryPkgVar(PERL_OBJECT_THIS_ stash,
				    LCID_NAME, LCID_LEN, lcidDefault);

	    hr = VariantChangeTypeEx(&Variant, pVariant, lcid, 0, VT_BSTR);
	    pVariant = &Variant;
	}

	if (!CheckOleError(PERL_OBJECT_THIS_ stash, hr)) {
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
