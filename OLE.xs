/* OLE.xs
 *
 *  (c) 1995 Microsoft Corporation. All rights reserved.
 *  Developed by ActiveWare Internet Corp., now known as
 *  ActiveState Tool Corp., http://www.ActiveState.com
 *
 *  Other modifications Copyright (c) 1997-1999 by Gurusamy Sarathy
 *  <gsar@activestate.com> and Jan Dubois <jand@activestate.com>
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

// #define _DEBUG

#define MY_VERSION "Win32::OLE(" XS_VERSION ")"

#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#define _WIN32_DCOM
#include <windows.h>
#include <ocidl.h>

#ifdef _DEBUG
#   include <crtdbg.h>
#   define DEBUGBREAK _CrtDbgBreak()
#else
#   define DEBUGBREAK
#endif

#if defined (__cplusplus)
extern "C" {
#endif

#ifdef __CYGWIN__
#   undef WIN32			/* don't use with Cygwin & Perl */
#   include <netdb.h>
#   include <sys/socket.h>
#   include <unistd.h>
    char *_strrev(char *);	/* from string.h (msvcrt40) */
#endif

#define MIN_PERL_DEFINE
#define NO_XSLOCKS
#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"
#include "patchlevel.h"

#undef WORD
typedef unsigned short WORD;

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

#ifndef pTHX_
#   define pTHX_
#endif

#undef THIS_
#define THIS_ PERL_OBJECT_THIS_

#if !defined(_DEBUG)
#   define DBG(a)
#else
#   define DBG(a)  MyDebug a
void
MyDebug(const char *pat, ...)
{
    char szBuffer[512];
    va_list args;
    va_start(args, pat);
    vsprintf(szBuffer, pat, args);
    OutputDebugString(szBuffer);
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
#define COINIT_NO_INITIALIZE -2

typedef HRESULT (STDAPICALLTYPE FNCOINITIALIZEEX)(LPVOID, DWORD);
typedef void (STDAPICALLTYPE FNCOUNINITIALIZE)(void);
typedef HRESULT (STDAPICALLTYPE FNCOCREATEINSTANCEEX)
    (REFCLSID, IUnknown*, DWORD, COSERVERINFO*, DWORD, MULTI_QI*);

typedef HWND (WINAPI FNHTMLHELP)(HWND hwndCaller, LPCSTR pszFile,
				 UINT uCommand, DWORD dwData);

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

    /* HTML Help Control loaded dynamically */
    HINSTANCE hHHCTRL;
    FNHTMLHELP *pfnHtmlHelp;

}   PERINTERP;

#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
#   if (PATCHLEVEL == 4) && (SUBVERSION < 68)
#       define dPERINTERP                                                 \
           SV *interp = perl_get_sv(MY_VERSION, FALSE);                   \
           if (!interp || !SvIOK(interp))	                          \
               warn(MY_VERSION ": Per-interpreter data not initialized"); \
           PERINTERP *pInterp = (PERINTERP*)SvIV(interp)
#   else
#	define dPERINTERP                                                 \
           SV **pinterp = hv_fetch(PL_modglobal, MY_VERSION,              \
                                   sizeof(MY_VERSION)-1, FALSE);          \
           if (!pinterp || !*pinterp || !SvIOK(*pinterp))		  \
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

#define g_hHHCTRL               (INTERP->hHHCTRL)
#define g_pfnHtmlHelp           (INTERP->pfnHtmlHelp)

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
class EventSink;
typedef struct
{
    OBJECTHEADER header;

    BOOL bDestroyed;
    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo;
    IEnumVARIANT *pEnum;
    EventSink *pEventSink;

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

/* EventSink class */
class EventSink : public IDispatch
{
 public:
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, LPVOID *ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);

    // IDispatch methods
    STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
    STDMETHOD(GetTypeInfo)(
      UINT itinfo,
      LCID lcid,
      ITypeInfo **pptinfo);
    STDMETHOD(GetIDsOfNames)(
      REFIID riid,
      OLECHAR **rgszNames,
      UINT cNames,
      LCID lcid,
      DISPID *rgdispid);
    STDMETHOD(Invoke)(
      DISPID dispidMember,
      REFIID riid,
      LCID lcid,
      WORD wFlags,
      DISPPARAMS *pdispparams,
      VARIANT *pvarResult,
      EXCEPINFO *pexcepinfo,
      UINT *puArgErr);

#ifdef _DEBUG
    STDMETHOD(Dummy1)();
    STDMETHOD(Dummy2)();
    STDMETHOD(Dummy3)();
    STDMETHOD(Dummy4)();
    STDMETHOD(Dummy5)();
    STDMETHOD(Dummy6)();
    STDMETHOD(Dummy7)();
    STDMETHOD(Dummy8)();
    STDMETHOD(Dummy9)();
    STDMETHOD(Dummy10)();
    STDMETHOD(Dummy11)();
    STDMETHOD(Dummy12)();
    STDMETHOD(Dummy13)();
    STDMETHOD(Dummy14)();
    STDMETHOD(Dummy15)();
    STDMETHOD(Dummy16)();
    STDMETHOD(Dummy17)();
    STDMETHOD(Dummy18)();
    STDMETHOD(Dummy19)();
    STDMETHOD(Dummy20)();
    STDMETHOD(Dummy21)();
    STDMETHOD(Dummy22)();
    STDMETHOD(Dummy23)();
    STDMETHOD(Dummy24)();
    STDMETHOD(Dummy25)();
#endif

    EventSink(CPERLarg_ WINOLEOBJECT *pObj, SV *events,
	      REFIID riid, ITypeInfo *pTypeInfo);
    ~EventSink(void);
    HRESULT Advise(IConnectionPoint *pConnectionPoint);
    void Unadvise(void);

 private:
    int m_refcount;
    WINOLEOBJECT *m_pObj;
    IConnectionPoint *m_pConnectionPoint;
    DWORD m_dwCookie;

    SV *m_events;
    IID m_iid;
    ITypeInfo *m_pTypeInfo;
#ifdef PERL_OBJECT
    CPERLproto m_PERL_OBJECT_THIS;
#endif
};

/* Forwarder class */
class Forwarder : public IDispatch
{
 public:
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, LPVOID *ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);

    // IDispatch methods
    STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
    STDMETHOD(GetTypeInfo)(
      UINT itinfo,
      LCID lcid,
      ITypeInfo **pptinfo);
    STDMETHOD(GetIDsOfNames)(
      REFIID riid,
      OLECHAR **rgszNames,
      UINT cNames,
      LCID lcid,
      DISPID *rgdispid);
    STDMETHOD(Invoke)(
      DISPID dispidMember,
      REFIID riid,
      LCID lcid,
      WORD wFlags,
      DISPPARAMS *pdispparams,
      VARIANT *pvarResult,
      EXCEPINFO *pexcepinfo,
      UINT *puArgErr);

    Forwarder(CPERLarg_ HV *stash, SV *method);
    ~Forwarder(void);

 private:
    int m_refcount;
    HV *m_stash;
    SV *m_method;
#ifdef PERL_OBJECT
    CPERLproto m_PERL_OBJECT_THIS;
#endif
};

/* forward declarations */
HRESULT SetSVFromVariantEx(CPERLarg_ VARIANTARG *pVariant, SV* sv, HV *stash,
			   BOOL bByRefObj=FALSE);
HRESULT SetVariantFromSVEx(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp,
			   LCID lcid);
HRESULT AssignVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant,
			    UINT cp, LCID lcid);

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

/* SvPV_nolen() macro first defined in 5.005_55 */
#if (PATCHLEVEL == 4) || ((PATCHLEVEL == 5) && (SUBVERSION < 55))
char *
MySvPVX(CPERLarg_ SV *sv)
{
    STRLEN n_a;
    return SvPV(sv, n_a);
}
#    define SvPV_nolen(sv) (SvPOK(sv) ? (SvPVX(sv)) : MySvPVX(THIS_ sv))
#endif

//------------------------------------------------------------------------

inline void
SpinMessageLoop(void)
{
    MSG msg;

    DBG(("SpinMessageLoop\n"));
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
	TranslateMessage(&msg);
	DispatchMessage(&msg);
    }

}   /* SpinMessageLoop */

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
    if (!pHostEnt)
	return FALSE;

    if (pHostEnt->h_addrtype != PF_INET || pHostEnt->h_length != 4) {
	warn(MY_VERSION ": IsLocalMachine() gethostbyname failure");
	return FALSE;
    }

    int index;
    int count = 0;
    char *pLocal;
    while (pHostEnt->h_addr_list[count])
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
	if (pHostEnt)
	    if (pHostEnt->h_addrtype == PF_INET && pHostEnt->h_length == 4)
		ppRemote = pHostEnt->h_addr_list;
    }

    /* Compare list of addresses of remote machine against local addresses */
    while (*ppRemote) {
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

    err = RegConnectRegistry(pszHost, HKEY_LOCAL_MACHINE, &hKeyLocalMachine);
    if (err != ERROR_SUCCESS)
	return HRESULT_FROM_WIN32(err);

    SV *subkey = sv_2mortal(newSVpv("SOFTWARE\\Classes\\", 0));
    sv_catpv(subkey, pszProgID);
    sv_catpv(subkey, "\\CLSID");

    err = RegOpenKeyEx(hKeyLocalMachine, SvPV_nolen(subkey), 0, KEY_READ,
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
    if (pszHeap != pszStack && pszHeap)
	Safefree(pszHeap);
}

char *
GetMultiByte(CPERLarg_ OLECHAR *wide, char *psz, int len, UINT cp)
{
    int count;

    if (psz) {
	if (!wide) {
	    *psz = (char) 0;
	    return psz;
	}
	count = WideCharToMultiByte(cp, 0, wide, -1, psz, len, NULL, NULL);
	if (count > 0)
	    return psz;
    }
    else if (!wide) {
	Newz(0, psz, 1, char);
	return psz;
    }

    count = WideCharToMultiByte(cp, 0, wide, -1, NULL, 0, NULL, NULL);
    if (count == 0) { /* should never happen */
	warn(MY_VERSION ": GetMultiByte() failure: %lu", GetLastError());
	DEBUGBREAK;
	if (!psz)
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

    pszBuffer = GetMultiByte(THIS_ wide, szBuffer, sizeof(szBuffer), cp);
    if (!sv)
	sv = newSVpv(pszBuffer, 0);
    else
	sv_setpv(sv, pszBuffer);
    ReleaseBuffer(THIS_ pszBuffer, szBuffer);
    return sv;
}

OLECHAR *
GetWideChar(CPERLarg_ char *psz, OLECHAR *wide, int len, UINT cp)
{
    /* Note: len is number of OLECHARs, not bytes! */
    int count;

    if (wide) {
	if (!psz) {
	    *wide = (OLECHAR) 0;
	    return wide;
	}
	count = MultiByteToWideChar(cp, 0, psz, -1, wide, len);
	if (count > 0)
	    return wide;
    }
    else if (!psz) {
	Newz(0, wide, 1, OLECHAR);
	return wide;
    }

    count = MultiByteToWideChar(cp, 0, psz, -1, NULL, 0);
    if (count == 0) {
	warn(MY_VERSION ": GetWideChar() failure: %lu", GetLastError());
	DEBUGBREAK;
	if (!wide)
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
	return (HV*)&PL_sv_undef;

}   /* GetStash */

HV *
GetWin32OleStash(CPERLarg_ SV *sv)
{
    SV *pkg;

    if (sv_isobject(sv))
	pkg = newSVpv(HvNAME(SvSTASH(SvRV(sv))), 0);
    else if (SvPOK(sv))
	pkg = newSVpv(SvPVX(sv), SvCUR(sv));
    else
	pkg = newSVpv(szWINOLE, 0); /* should never happen */

    char *pszColon = strrchr(SvPVX(pkg), ':');
    if (pszColon) {
	--pszColon;
	while (pszColon > SvPVX(pkg) && *pszColon == ':')
	    --pszColon;
	SvCUR_set(pkg, pszColon - SvPVX(pkg) + 1);
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

    if (gv && (sv = GvSV(*gv)) != NULL && SvIOK(sv)) {
	DBG(("QueryPkgVar(%s) returns %d\n", var, SvIV(sv)));
	return SvIV(sv);
    }

    return def;
}

void
SetLastOleError(CPERLarg_ HV *stash, HRESULT hr=S_OK, char *pszMsg=NULL)
{
    /* Find $Win32::OLE::LastError */
    SV *sv = sv_2mortal(newSVpv(HvNAME(stash), 0));
    sv_catpvn(sv, "::", 2);
    sv_catpvn(sv, LASTERR_NAME, LASTERR_LEN);
    SV *lasterr = perl_get_sv(SvPV_nolen(sv), TRUE);
    if (!lasterr) {
	warn(MY_VERSION ": SetLastOleError: couldnot create variable %s",
	     LASTERR_NAME);
	DEBUGBREAK;
	return;
    }

    sv_setiv(lasterr, (IV)hr);
    if (pszMsg) {
	sv_setpv(lasterr, pszMsg);
	SvIOK_on(lasterr);
    }
}

void
ReportOleError(CPERLarg_ HV *stash, HRESULT hr, EXCEPINFO *pExcep=NULL,
	       SV *svAdd=NULL)
{
    dSP;

    SV *sv;
    IV warnlvl = QueryPkgVar(THIS_ stash, WARN_NAME, WARN_LEN);
    GV **pgv = (GV**)hv_fetch(stash, WARN_NAME, WARN_LEN, FALSE);
    CV *cv = Nullcv;

    if (pgv && (sv = GvSV(*pgv)) && SvROK(sv) && SvTYPE(SvRV(sv)) == SVt_PVCV)
	cv = (CV*)sv;

    sv = sv_2mortal(newSV(200));
    SvPOK_on(sv);

    /* start with exception info */
    if (pExcep && (pExcep->bstrSource || pExcep->bstrDescription)) {
	char szSource[80] = "<Unknown Source>";
	char szDesc[200] = "<No description provided>";

	char *pszSource = szSource;
	char *pszDesc = szDesc;

	UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);
	if (pExcep->bstrSource)
	    pszSource = GetMultiByte(THIS_ pExcep->bstrSource,
				     szSource, sizeof(szSource), cp);

	if (pExcep->bstrDescription)
	    pszDesc = GetMultiByte(THIS_ pExcep->bstrDescription,
				   szDesc, sizeof(szDesc), cp);

	sv_setpvf(sv, "OLE exception from \"%s\":\n\n%s\n\n",
		  pszSource, pszDesc);

	ReleaseBuffer(THIS_ pszSource, szSource);
	ReleaseBuffer(THIS_ pszDesc, szDesc);
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
    if (svAdd) {
	sv_catpv(sv, "\n    ");
	sv_catsv(sv, svAdd);
    }

    /* try to keep linelength of description below 80 chars. */
    char *pLastBlank = NULL;
    char *pch = SvPVX(sv);
    int  cch;

    for (cch = 0 ; *pch ; ++pch, ++cch) {
	if (*pch == ' ') {
	    pLastBlank = pch;
	}
	else if (*pch == '\n') {
	    pLastBlank = pch;
	    cch = 0;
	}

	if (cch > 76 && pLastBlank) {
	    *pLastBlank = '\n';
	    cch = pch - pLastBlank;
	}
    }

    SetLastOleError(THIS_ stash, hr, SvPVX(sv));

    DBG(("ReportOleError: hr=0x%08x warnlvl=%d\n%s", hr, warnlvl, SvPVX(sv)));

    if (!cv && (warnlvl > 1 || (warnlvl == 1 && PL_dowarn))) {
	if (warnlvl < 3) {
	    cv = perl_get_cv("Carp::carp", FALSE);
	    if (!cv)
		warn(SvPVX(sv));
	}
	else {
	    cv = perl_get_cv("Carp::croak", FALSE);
	    if (!cv)
		croak(SvPVX(sv));
	}
    }

    if (cv) {
        PUSHMARK(sp) ;
        XPUSHs(sv);
        PUTBACK;
        perl_call_sv((SV*)cv, G_DISCARD);
    }

}   /* ReportOleError */

inline BOOL
CheckOleError(CPERLarg_ HV *stash, HRESULT hr, EXCEPINFO *pExcep=NULL,
	      SV *svAdd=NULL)
{
    if (FAILED(hr)) {
	ReportOleError(THIS_ stash, hr, pExcep, svAdd);
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
    DBG(("AddToObjectChain(0x%08x) lMagic=0x%08x", pHeader, lMagic));

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
    DBG(("RemoveFromObjectChain(0x%08x) lMagic=0x%08x\n", pHeader,
	 pHeader ? pHeader->lMagic : 0));

    if (!pHeader)
	return;

#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    PERINTERP *pInterp = pHeader->pInterp;
#endif

    EnterCriticalSection(&g_CriticalSection);
    if (!pHeader->pPrevious) {
	g_pObj = pHeader->pNext;
	if (g_pObj)
	    g_pObj->pPrevious = NULL;
    }
    else if (!pHeader->pNext)
	pHeader->pPrevious->pNext = NULL;
    else {
	pHeader->pPrevious->pNext = pHeader->pNext;
	pHeader->pNext->pPrevious = pHeader->pPrevious;
    }
    pHeader->lMagic = 0;
    LeaveCriticalSection(&g_CriticalSection);
}

SV *
CreatePerlObject(CPERLarg_ HV *stash, IDispatch *pDispatch, SV *destroy)
{
    /* returns a mortal reference to a new Perl OLE object */

    if (!pDispatch) {
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

    if (gv && (sv = GvSV(*gv)) != NULL && SvPOK(sv))
	szTie = SvPV_nolen(sv);

    New(0, pObj, 1, WINOLEOBJECT);
    pObj->bDestroyed = FALSE;
    pObj->pDispatch = pDispatch;
    pObj->pTypeInfo = NULL;
    pObj->pEnum = NULL;
    pObj->pEventSink = NULL;
    pObj->hashTable = newHV();
    pObj->self = newHV();

    pObj->destroy = NULL;
    if (destroy) {
	if (SvPOK(destroy))
	    pObj->destroy = newSVsv(destroy);
	else if (SvROK(destroy) && SvTYPE(SvRV(destroy)) == SVt_PVCV)
	    pObj->destroy = newRV_inc(SvRV(destroy));
    }

    AddToObjectChain(THIS_ &pObj->header, WINOLE_MAGIC);

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

    DBG(("ReleasePerlObject |%lx|", pObj));

    if (!pObj)
	return;

    /* ReleasePerlObject may be called multiple times for a single object:
     * first by Uninitialize() and then by Win32::OLE::DESTROY.
     * Make sure nothing is cleaned up twice!
     */

    if (pObj->destroy) {
	SV *self = sv_2mortal(newRV_inc((SV*)pObj->self));

	/* honour OVERLOAD setting */
	if (Gv_AMG(SvSTASH(pObj->self)))
	    SvAMAGIC_on(self);

	DBG(("Calling destroy method for object |%lx|\n", pObj));
	ENTER;
	if (SvPOK(pObj->destroy)) {
	    /* $self->Dispatch($destroy,$retval); */
	    EXTEND(sp, 3);
	    PUSHMARK(sp);
	    PUSHs(self);
	    PUSHs(pObj->destroy);
	    PUSHs(sv_newmortal());
	    PUTBACK;
	    perl_call_method("Dispatch", G_DISCARD);
	}
	else {
	    /* &$destroy($self); */
	    PUSHMARK(sp);
	    XPUSHs(self) ;
	    PUTBACK;
	    perl_call_sv(pObj->destroy, G_DISCARD);
	}
	LEAVE;
	DBG(("Returned from destroy method for 0x%08x\n", pObj));

	SvREFCNT_dec(pObj->destroy);
	pObj->destroy = NULL;
    }

    if (pObj->pEventSink) {
	DBG(("Unadvise connection |%lx|", pObj));
	pObj->pEventSink->Unadvise();
	pObj->pEventSink = NULL;
    }

    if (pObj->pDispatch) {
	DBG((" pDispatch"));
	pObj->pDispatch->Release();
	pObj->pDispatch = NULL;
    }

    if (pObj->pTypeInfo) {
	DBG((" pTypeInfo"));
	pObj->pTypeInfo->Release();
	pObj->pTypeInfo = NULL;
    }

    if (pObj->pEnum) {
	DBG((" pEnum"));
	pObj->pEnum->Release();
	pObj->pEnum = NULL;
    }

    if (pObj->destroy) {
	DBG((" destroy(%d)", SvREFCNT(pObj->destroy)));
	SvREFCNT_dec(pObj->destroy);
	pObj->destroy = NULL;
    }

    if (pObj->hashTable) {
	DBG((" hashTable(%d)", SvREFCNT(pObj->hashTable)));
	SvREFCNT_dec(pObj->hashTable);
	pObj->hashTable = NULL;
    }

    DBG(("\n"));

}   /* ReleasePerlObject */

WINOLEOBJECT *
GetOleObject(CPERLarg_ SV *sv, BOOL bDESTROY=FALSE)
{
    if (sv_isobject(sv) && SvTYPE(SvRV(sv)) == SVt_PVHV) {
	SV **psv = hv_fetch((HV*)SvRV(sv), PERL_OLE_ID, PERL_OLE_IDLEN, 0);

	/* Win32::OLE::Tie::DESTROY called before Win32::OLE::DESTROY? */
	if (!psv && bDESTROY)
	    return NULL;

#if (PATCHLEVEL > 4) || ((PATCHLEVEL == 4) && (SUBVERSION > 4))
	if (psv && SvGMAGICAL(*psv))
	    mg_get(*psv);

	if (psv && SvIOK(*psv)) {
#else
	if (psv) {
#endif
	    WINOLEOBJECT *pObj = (WINOLEOBJECT*)SvIV(*psv);

	    DBG(("GetOleObject = |%lx|\n", pObj));
	    if (pObj && pObj->header.lMagic == WINOLE_MAGIC)
		if (pObj->pDispatch || bDESTROY)
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

	if (pEnumObj && pEnumObj->header.lMagic == WINOLEENUM_MAGIC)
	    if (pEnumObj->pEnum || bDESTROY)
		return pEnumObj;
    }
    warn(MY_VERSION ": GetOleEnumObject() Not a %s object", szWINOLEENUM);
    DEBUGBREAK;
    return (WINOLEENUMOBJECT*)NULL;
}

WINOLEVARIANTOBJECT *
GetOleVariantObject(CPERLarg_ SV *sv, BOOL bWarn=TRUE)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLEVARIANT)) {
	WINOLEVARIANTOBJECT *pVarObj = (WINOLEVARIANTOBJECT*)SvIV(SvRV(sv));

	if (pVarObj && pVarObj->header.lMagic == WINOLEVARIANT_MAGIC)
	    return pVarObj;
    }
    if (bWarn) {
	warn(MY_VERSION ": GetOleVariantObject() Not a %s object",
	     szWINOLEVARIANT);
	DEBUGBREAK;
    }
    return (WINOLEVARIANTOBJECT*)NULL;
}

SV *
CreateTypeLibObject(CPERLarg_ ITypeLib *pTypeLib, TLIBATTR *pTLibAttr)
{
    WINOLETYPELIBOBJECT *pObj;
    New(0, pObj, 1, WINOLETYPELIBOBJECT);

    pObj->pTypeLib = pTypeLib;
    pObj->pTLibAttr = pTLibAttr;

    AddToObjectChain(THIS_ (OBJECTHEADER*)pObj, WINOLETYPELIB_MAGIC);

    return sv_bless(newRV_noinc(newSViv((IV)pObj)),
		    gv_stashpv(szWINOLETYPELIB, TRUE));
}

WINOLETYPELIBOBJECT *
GetOleTypeLibObject(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLETYPELIB)) {
	WINOLETYPELIBOBJECT *pObj = (WINOLETYPELIBOBJECT*)SvIV(SvRV(sv));

	if (pObj && pObj->header.lMagic == WINOLETYPELIB_MAGIC)
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

    AddToObjectChain(THIS_ (OBJECTHEADER*)pObj, WINOLETYPEINFO_MAGIC);

    return sv_bless(newRV_noinc(newSViv((IV)pObj)),
		    gv_stashpv(szWINOLETYPEINFO, TRUE));
}

WINOLETYPEINFOOBJECT *
GetOleTypeInfoObject(CPERLarg_ SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLETYPEINFO)) {
	WINOLETYPEINFOOBJECT *pObj = (WINOLETYPEINFOOBJECT*)SvIV(SvRV(sv));

	if (pObj && pObj->header.lMagic == WINOLETYPEINFO_MAGIC)
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
    if (psv) {
	dispID = (DISPID)SvIV(*psv);
	return S_OK;
    }

    /* not there so get info and add it */
    DISPID id;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;

    pBuffer = GetWideChar(THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    hr = pObj->pDispatch->GetIDsOfNames(IID_NULL, &pBuffer, 1, lcid, &id);
    ReleaseBuffer(THIS_ pBuffer, Buffer);
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

    if (pObj->pTypeInfo)
	return;

    HRESULT hr = pObj->pDispatch->GetTypeInfoCount(&count);
    if (hr == E_NOTIMPL || count == 0) {
	DBG(("GetTypeInfoCount returned %u (count=%d)", hr, count));
	return;
    }

    if (CheckOleError(THIS_ stash, hr)) {
	warn(MY_VERSION ": FetchTypeInfo() GetTypeInfoCount failed");
	DEBUGBREAK;
	return;
    }

    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(THIS_ stash, hr))
	return;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ stash, hr);
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
	ReportOleError(THIS_ stash, hr);
	return;
    }

    if (pTypeAttr) {
	if (pTypeAttr->typekind == TKIND_DISPATCH) {
	    pObj->cFuncs = pTypeAttr->cFuncs;
	    pObj->cVars = pTypeAttr->cVars;
	    pObj->PropIndex = 0;
	    pObj->pTypeInfo = pTypeInfo;
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	if (!pObj->pTypeInfo)
	    pTypeInfo->Release();
    }

}   /* FetchTypeInfo */

SV *
NextPropertyName(CPERLarg_ WINOLEOBJECT *pObj)
{
    HRESULT hr;
    unsigned int cName;
    BSTR bstr;

    if (!pObj->pTypeInfo)
	return &PL_sv_undef;

    HV *stash = SvSTASH(pObj->self);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    while (pObj->PropIndex < pObj->cFuncs+pObj->cVars) {
	ULONG index = pObj->PropIndex++;
	/* Try all the INVOKE_PROPERTYGET functions first */
	if (index < pObj->cFuncs) {
	    FUNCDESC *pFuncDesc;

	    hr = pObj->pTypeInfo->GetFuncDesc(index, &pFuncDesc);
	    if (CheckOleError(THIS_ stash, hr))
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
	    if (CheckOleError(THIS_ stash, hr) || cName == 0 || !bstr)
		continue;

	    SV *sv = sv_setwide(THIS_ NULL, bstr, cp);
	    SysFreeString(bstr);
	    return sv;
	}
	/* Now try the VAR_DISPATCH kind variables used by older OLE versions */
	else {
	    VARDESC *pVarDesc;

	    index -= pObj->cFuncs;
	    hr = pObj->pTypeInfo->GetVarDesc(index, &pVarDesc);
	    if (CheckOleError(THIS_ stash, hr))
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
	    if (CheckOleError(THIS_ stash, hr) || cName == 0 || !bstr)
		continue;

	    SV *sv = sv_setwide(THIS_ NULL, bstr, cp);
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

    pszStr = GetMultiByte(THIS_ bstrName, szStr, sizeof(szStr), cp);
    hv_store(hv, "Name", 4, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(THIS_ pszStr, szStr);
    SysFreeString(bstrName);

    pszStr = GetMultiByte(THIS_ bstrDocString, szStr, sizeof(szStr), cp);
    hv_store(hv, "DocString", 9, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(THIS_ pszStr, szStr);
    SysFreeString(bstrDocString);

    pszStr = GetMultiByte(THIS_ bstrHelpFile, szStr, sizeof(szStr), cp);
    hv_store(hv, "HelpFile", 8, newSVpv(pszStr, 0), 0);
    ReleaseBuffer(THIS_ pszStr, szStr);
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
	    if (SUCCEEDED(hr))
		sv = CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr);
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
	hr = TranslateTypeDesc(THIS_ pTypeDesc->lptdesc, pObj, av);

    return hr;
}

HV *
TranslateElemDesc(CPERLarg_ ELEMDESC *pElemDesc, WINOLETYPEINFOOBJECT *pObj,
		  HV *olestash)
{
    HV *hv = newHV();

    AV *av = newAV();
    TranslateTypeDesc(THIS_  &pElemDesc->tdesc, pObj, av);
    hv_store(hv, "vt", 2, newRV_noinc((SV*)av), 0);

    USHORT wParamFlags = pElemDesc->paramdesc.wParamFlags;
    hv_store(hv, "wParamFlags", 11, newSViv(wParamFlags), 0);

    USHORT wMask = PARAMFLAG_FOPT|PARAMFLAG_FHASDEFAULT;
    if ((wParamFlags & wMask) == wMask) {
	PARAMDESCEX *pParamDescEx = pElemDesc->paramdesc.pparamdescex;
	hv_store(hv, "cBytes", 6, newSViv(pParamDescEx->cBytes), 0);
	// XXX should be stored as a Win32::OLE::Variant object ?
	SV *sv = newSV(0);
	// XXX check return code
	SetSVFromVariantEx(THIS_ &pParamDescEx->varDefaultValue,
			   sv, olestash);
	hv_store(hv, "varDefaultValue", 15, sv, 0);
    }

    return hv;

}   /* TranslateElemDesc */

HRESULT
FindIID(CPERLarg_ WINOLEOBJECT *pObj, char *pszItf, IID *piid,
	ITypeInfo **ppTypeInfo, UINT cp, LCID lcid)
{
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;

    if (ppTypeInfo)
	*ppTypeInfo = NULL;

    // Determine containing type library
    HRESULT hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    DBG(("  GetTypeInfo: 0x%08x\n", hr));
    if (FAILED(hr))
	return hr;

    unsigned int index;
    hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
    pTypeInfo->Release();
    DBG(("  GetContainingTypeLib: 0x%08x\n", hr));
    if (FAILED(hr))
	return hr;

    // piid maybe already set by IProvideClassInfo2::GetGUID
    if (!pszItf) {
	hr = pTypeLib->GetTypeInfoOfGuid(*piid, ppTypeInfo);
	DBG(("  GetTypeInfoOfGuid: 0x%08x\n", hr));
	pTypeLib->Release();
	return hr;
    }

    // Walk through all type definitions in the library
    BOOL bFound = FALSE;
    unsigned int count = pTypeLib->GetTypeInfoCount();
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

	// DBG(("  TypeInfo %d typekind %d\n", index, pTypeAttr->typekind));

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
		if (FAILED(hr)) {
		    pImplTypeInfo->Release();
		    break;
		}

		char szStr[OLE_BUF_SIZ];
		char *pszStr = GetMultiByte(THIS_ bstr, szStr,
					    sizeof(szStr), cp);
		if (strEQ(pszItf, pszStr)) {
		    TYPEATTR *pImplTypeAttr;

		    hr = pImplTypeInfo->GetTypeAttr(&pImplTypeAttr);
		    if (SUCCEEDED(hr)) {
			bFound = TRUE;
			*piid = pImplTypeAttr->guid;
			if (ppTypeInfo) {
			    *ppTypeInfo = pImplTypeInfo;
			    (*ppTypeInfo)->AddRef();
			}
			pImplTypeInfo->ReleaseTypeAttr(pImplTypeAttr);
		    }
		}

		ReleaseBuffer(THIS_ pszStr, szStr);
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
    DBG(("  after loop: 0x%08x\n", hr));
    if (FAILED(hr))
	return hr;

    if (!bFound) {
	warn(MY_VERSION "FindIID: Interface '%s' not found", pszItf);
	return E_NOINTERFACE;
    }

#ifdef _DEBUG
    OLECHAR wszGUID[80];
    int len = StringFromGUID2(*piid, wszGUID, sizeof(wszGUID)/sizeof(OLECHAR));
    char szStr[OLE_BUF_SIZ];
    char *pszStr = GetMultiByte(THIS_ wszGUID, szStr, sizeof(szStr), cp);
    DBG(("FindIID: %s is %s", pszItf, pszStr));
    ReleaseBuffer(THIS_ pszStr, szStr);
#endif

    return S_OK;

}   /* FindIID */

HRESULT
FindDefaultSource(CPERLarg_ WINOLEOBJECT *pObj, IID *piid,
		  ITypeInfo **ppTypeInfo, UINT cp, LCID lcid)
{
    HRESULT hr;
    *ppTypeInfo = NULL;

    // Try IProvideClassInfo2 interface first
    IProvideClassInfo2 *pProvideClassInfo2;
    hr = pObj->pDispatch->QueryInterface(IID_IProvideClassInfo2,
					 (void**)&pProvideClassInfo2);
    DBG(("QueryInterface(IProvideClassInfo2): hr=0x%08x\n", hr));
    if (SUCCEEDED(hr)) {
	hr = pProvideClassInfo2->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID,
					 piid);
	pProvideClassInfo2->Release();
	DBG(("GetGUID: hr=0x%08x\n", hr));
	return FindIID(THIS_ pObj, NULL, piid, ppTypeInfo, cp, lcid);
    }

    IProvideClassInfo *pProvideClassInfo;
    hr = pObj->pDispatch->QueryInterface(IID_IProvideClassInfo,
					 (void**)&pProvideClassInfo);
    DBG(("QueryInterface(IProvideClassInfo): hr=0x%08x\n", hr));
    if (FAILED(hr))
	return hr;

    // Get ITypeInfo* for COCLASS of this object
    ITypeInfo *pTypeInfo;
    hr = pProvideClassInfo->GetClassInfo(&pTypeInfo);
    pProvideClassInfo->Release();
    DBG(("GetClassInfo: hr=0x%08x\n", hr));
    if (FAILED(hr))
	return hr;

    // Get Type Attributes
    TYPEATTR *pTypeAttr;
    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    DBG(("GetTypeAttr: hr=0x%08x\n", hr));
    if (FAILED(hr)) {
	pTypeInfo->Release();
	return hr;
    }

    UINT i;
    int iFlags;

    // Enumerate all implemented types of the COCLASS
    for (i=0; i < pTypeAttr->cImplTypes; i++) {
	hr = pTypeInfo->GetImplTypeFlags(i, &iFlags);
	DBG(("GetImplTypeFlags: hr=0x%08x i=%d iFlags=%d\n", hr, i, iFlags));
	if (FAILED(hr))
	    continue;

	// looking for the [default] [source]
	// we just hope that it is a dispinterface :-)
	if ((iFlags & IMPLTYPEFLAG_FDEFAULT) &&
	    (iFlags & IMPLTYPEFLAG_FSOURCE))
	{
	    HREFTYPE hRefType = NULL;

	    hr = pTypeInfo->GetRefTypeOfImplType(i, &hRefType);
	    DBG(("GetRefTypeOfImplType: hr=0x%08x\n", hr));
	    if (FAILED(hr))
		continue;
	    hr = pTypeInfo->GetRefTypeInfo(hRefType, ppTypeInfo);
	    DBG(("GetRefTypeInfo: hr=0x%08x\n", hr));
	    if (SUCCEEDED(hr))
		break;
	}
    }

    pTypeInfo->ReleaseTypeAttr(pTypeAttr);
    pTypeInfo->Release();

    // Now that would be a bad surprise, if we didn't find it, wouldn't it?
    if (!*ppTypeInfo) {
	if (SUCCEEDED(hr))
	    hr = E_UNEXPECTED;
	return hr;
    }

    // Determine IID of default source interface
    hr = (*ppTypeInfo)->GetTypeAttr(&pTypeAttr);
    if (SUCCEEDED(hr)) {
	*piid = pTypeAttr->guid;
	(*ppTypeInfo)->ReleaseTypeAttr(pTypeAttr);
    }
    else
	(*ppTypeInfo)->Release();

    return hr;

}   /* FindDefaultSource */

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
    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);

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
    CheckOleError(THIS_ stash, hr, &excepinfo);
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
	sv = newSV(0);
	hr = SetSVFromVariantEx(THIS_ &variant, sv, stash);
    }
    VariantClear(&variant);
    if (FAILED(hr)) {
        SvREFCNT_dec(sv);
	sv = &PL_sv_undef;
	ReportOleError(THIS_ stash, hr);
    }
    return sv;

}   /* NextEnumElement */

//------------------------------------------------------------------------

EventSink::EventSink(CPERLarg_ WINOLEOBJECT *pObj, SV *events,
		     REFIID riid, ITypeInfo *pTypeInfo)
{
    DBG(("EventSink::EventSink\n"));
    m_pObj = pObj;
    m_events = newSVsv(events);
    m_iid = riid;
    m_pTypeInfo = pTypeInfo;
    m_refcount = 1;
#ifdef PERL_OBJECT
    m_PERL_OBJECT_THIS = PERL_OBJECT_THIS;
#endif
}

EventSink::~EventSink(void)
{
#ifdef PERL_OBJECT
    CPERLarg = m_PERL_OBJECT_THIS;
#endif
    DBG(("EventSink::~EventSink\n"));
    if (m_pTypeInfo)
	m_pTypeInfo->Release();
    SvREFCNT_dec(m_events);
}

HRESULT
EventSink::Advise(IConnectionPoint *pConnectionPoint)
{
    HRESULT hr = pConnectionPoint->Advise((IUnknown*)this, &m_dwCookie);
    if (SUCCEEDED(hr)) {
	m_pConnectionPoint = pConnectionPoint;
	m_pConnectionPoint->AddRef();
    }
    return hr;
}

void
EventSink::Unadvise(void)
{
    if (m_pConnectionPoint) {
	m_pConnectionPoint->Unadvise(m_dwCookie);
	m_pConnectionPoint->Release();
    }
    m_pConnectionPoint = NULL;
    Release();
}

STDMETHODIMP
EventSink::QueryInterface(REFIID iid, void **ppv)
{
#ifdef _DEBUG
#   ifdef PERL_OBJECT
    CPERLarg = m_PERL_OBJECT_THIS;
#   endif
    OLECHAR wszGUID[80];
    int len = StringFromGUID2(iid, wszGUID, sizeof(wszGUID)/sizeof(OLECHAR));
    char szStr[OLE_BUF_SIZ];
    char *pszStr = GetMultiByte(THIS_ wszGUID, szStr, sizeof(szStr), CP_ACP);
    DBG(("***QueryInterface %s\n", pszStr));
    ReleaseBuffer(THIS_ pszStr, szStr);
#endif

    if (iid == IID_IUnknown || iid == IID_IDispatch || iid == m_iid)
	*ppv = this;
    else {
	DBG(("  failed\n"));
	*ppv = NULL;
	return E_NOINTERFACE;
    }
    DBG(("  succeeded\n"));
    AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG)
EventSink::AddRef(void)
{
    ++m_refcount;
    DBG(("***AddRef refcount=%d\n", m_refcount));
    return m_refcount;
}

STDMETHODIMP_(ULONG)
EventSink::Release(void)
{
    --m_refcount;
    DBG(("***Release refcount=%d\n", m_refcount));
    if (m_refcount)
	return m_refcount;
    delete this;
    return 0;
}

STDMETHODIMP
EventSink::GetTypeInfoCount(UINT *pctinfo)
{
    DBG(("***GetTypeInfoCount\n"));
    *pctinfo = 0;
    return S_OK;
}

STDMETHODIMP
EventSink::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
    DBG(("***GetTypeInfo\n"));
    *pptinfo = NULL;
    return DISP_E_BADINDEX;
}

STDMETHODIMP
EventSink::GetIDsOfNames(
    REFIID riid,
    OLECHAR **rgszNames,
    UINT cNames,
    LCID lcid,
    DISPID *rgdispid)
{
    DBG(("***GetIDsOfNames\n"));
    // XXX Set all DISPIDs to DISPID_UNKNOWN
    return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP
EventSink::Invoke(
    DISPID dispidMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pdispparams,
    VARIANT *pvarResult,
    EXCEPINFO *pexcepinfo,
    UINT *puArgErr)
{
#ifdef PERL_OBJECT
    CPERLarg = m_PERL_OBJECT_THIS;
#endif

    DBG(("***Invoke dispid=%d args=%d\n", dispidMember, pdispparams->cArgs));
    BSTR bstr;
    unsigned int count;
    HRESULT hr;
    SV *event = Nullsv;

    if (m_pTypeInfo) {
	hr = m_pTypeInfo->GetNames(dispidMember, &bstr, 1, &count);
	if (FAILED(hr)) {
	    DBG(("  GetNames failed: 0x%08x\n", hr));
	    return S_OK;
	}

	event = sv_2mortal(sv_setwide(THIS_ NULL, bstr, CP_ACP));
	SysFreeString(bstr);
    }
    else {
	DBG(("  No type library available\n"));
	STRLEN n_a;
	event = sv_2mortal(newSViv(dispidMember));
	SvPV_force(event, n_a);
    }

    DBG(("  Event %s\n", SvPVX(event)));

    SV *callback = NULL;
    BOOL pushname = FALSE;

    if (SvROK(m_events) && SvTYPE(SvRV(m_events)) == SVt_PVCV) {
	callback = m_events;
	pushname = TRUE;
    }
    else if (SvPOK(m_events)) {
	HV *stash = gv_stashsv(m_events, FALSE);
	if (stash) {
	    GV **pgv = (GV**)hv_fetch(stash, SvPVX(event), SvCUR(event), FALSE);
	    if (pgv && GvCV(*pgv))
		callback = (SV*)GvCV(*pgv);
	}
    }

    if (callback) {
	dSP;
	SV *self = newRV_inc((SV*)m_pObj->self);
	if (Gv_AMG(SvSTASH(m_pObj->self)))
	    SvAMAGIC_on(self);

	ENTER ;
	SAVETMPS ;
	PUSHMARK(sp);
	XPUSHs(sv_2mortal(self));
	if (pushname)
	    XPUSHs(event);
	for (int i=0; i < pdispparams->cArgs; ++i) {
	    VARIANT *pVariant = &pdispparams->rgvarg[pdispparams->cArgs-i-1];
	    DBG(("   Arg %d vt=0x%04x\n", i, V_VT(pVariant)));
	    SV *sv = sv_newmortal();
	    // XXX Check return code
	    SetSVFromVariantEx(THIS_ pVariant, sv, SvSTASH(m_pObj->self), TRUE);
	    XPUSHs(sv);
	}
	PUTBACK;
	perl_call_sv(callback, G_DISCARD);
	SPAGAIN;
	FREETMPS ;
	LEAVE ;
    }

    return S_OK;
}

#ifdef _DEBUG
#define Dummy(i) STDMETHODIMP EventSink::Dummy##i(void) \
   {  DBG(("***Dummy%d\n", i)); return S_OK; }

Dummy(1)  Dummy(2)  Dummy(3)  Dummy(4)  Dummy(5)
Dummy(6)  Dummy(7)  Dummy(8)  Dummy(9)  Dummy(10)
Dummy(11) Dummy(12) Dummy(13) Dummy(14) Dummy(15)
Dummy(16) Dummy(17) Dummy(18) Dummy(19) Dummy(20)
Dummy(21) Dummy(22) Dummy(23) Dummy(24) Dummy(25)
#endif

//------------------------------------------------------------------------

Forwarder::Forwarder(CPERLarg_ HV *stash, SV *method)
{
    m_stash = stash; // XXX refcount?
    m_method = newSVsv(method);
    m_refcount = 1;
#ifdef PERL_OBJECT
    m_PERL_OBJECT_THIS = PERL_OBJECT_THIS;
#endif
}

Forwarder::~Forwarder(void)
{
#ifdef PERL_OBJECT
    CPERLarg = m_PERL_OBJECT_THIS;
#endif
    SvREFCNT_dec(m_method);
}

STDMETHODIMP
Forwarder::QueryInterface(REFIID iid, void **ppv)
{
    if (iid == IID_IUnknown || iid == IID_IDispatch) {
	*ppv = this;
	AddRef();
	return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG)
Forwarder::AddRef(void)
{
    return ++m_refcount;
}

STDMETHODIMP_(ULONG)
Forwarder::Release(void)
{
    if (--m_refcount)
	return m_refcount;
    delete this;
    return 0;
}

STDMETHODIMP
Forwarder::GetTypeInfoCount(UINT *pctinfo)
{
    *pctinfo = 0;
    return S_OK;
}

STDMETHODIMP
Forwarder::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
    *pptinfo = NULL;
    return DISP_E_BADINDEX;
}

STDMETHODIMP
Forwarder::GetIDsOfNames(
    REFIID riid,
    OLECHAR **rgszNames,
    UINT cNames,
    LCID lcid,
    DISPID *rgdispid)
{
    DBG(("Forwarder::GetIDsOfNames cNames=%d\n", cNames));
    // XXX Set all DISPIDs to DISPID_UNKNOWN
    return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP
Forwarder::Invoke(
    DISPID dispidMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pdispparams,
    VARIANT *pvarResult,
    EXCEPINFO *pexcepinfo,
    UINT *puArgErr)
{
#ifdef PERL_OBJECT
    CPERLarg = m_PERL_OBJECT_THIS;
#endif

    DBG(("Forwarder::Invoke dispid=%d args=%d\n",
	 dispidMember, pdispparams->cArgs));
    dSP;
    ENTER ;
    SAVETMPS ;
    PUSHMARK(sp);
    for (int i=0; i < pdispparams->cArgs; ++i) {
	VARIANT *pVariant = &pdispparams->rgvarg[pdispparams->cArgs-i-1];
	DBG(("   Arg %d vt=0x%04x\n", i, V_VT(pVariant)));
	SV *sv = sv_newmortal();
	// XXX Check return code
	SetSVFromVariantEx(THIS_ pVariant, sv, m_stash, TRUE);
	XPUSHs(sv);
    }
    PUTBACK;
    perl_call_sv(m_method, G_DISCARD);
    SPAGAIN;
    FREETMPS ;
    LEAVE ;
    return S_OK;
}

//------------------------------------------------------------------------

SV *
SetSVFromGUID(CPERLarg_ REFGUID rguid)
{
    dSP;
    SV *sv = newSVsv(&PL_sv_undef);
    CV *cv = perl_get_cv("Win32::COM::GUID::new", FALSE);

    if (cv) {
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
    else {
	OLECHAR wszGUID[80];
	int len = StringFromGUID2(rguid, wszGUID,
				  sizeof(wszGUID)/sizeof(OLECHAR));
	if (len > 0) {
	    wszGUID[len-2] = (OLECHAR) 0;
	    sv_setwide(THIS_ sv, wszGUID+1, CP_ACP);
	}
    }
    return sv;
}

HRESULT
SetSafeArrayFromAV(CPERLarg_ AV* av, VARTYPE vt, SAFEARRAY *psa,
		   UINT cDims, UINT cp, LCID lcid)
{
    HRESULT hr = SafeArrayLock(psa);
    if (FAILED(hr))
	return hr;

    if (cDims == 0)
	cDims = SafeArrayGetDim(psa);

    AV **pav;
    long *pix;
    long *plen;

    New(0, pav, cDims, AV*);
    New(0, pix, cDims, long);
    New(0, plen, cDims, long);

    pav[0] = av;
    plen[0] = av_len(pav[0])+1;
    Zero(pix, cDims, long);

    VARIANT variant;
    VARIANT *pElement = &variant;
    if (vt != VT_VARIANT)
	V_VT(pElement) = vt | VT_BYREF;

    for (int index = 0 ; index >= 0 ; ) {
	SV **psv = av_fetch(pav[index], pix[index], FALSE);

	if (psv) {
	    if (SvROK(*psv) && SvTYPE(SvRV(*psv)) == SVt_PVAV) {
		if (++index >= cDims) {
		    warn(MY_VERSION ": SetSafeArrayFromAV unexpected failure");
		    hr = E_UNEXPECTED;
		    break;
		}
		pav[index] = (AV*)SvRV(*psv);
		pix[index] = 0;
		plen[index] = av_len(pav[index])+1;
		continue;
	    }

	    if (SvOK(*psv)) {
		if (index+1 != cDims) {
		    warn(MY_VERSION ": SetSafeArrayFromAV wrong dimension");
		    hr = DISP_E_BADINDEX;
		    break;
		}
		if (vt == VT_VARIANT) {
		    hr = SafeArrayPtrOfIndex(psa, pix, (void**)&pElement);
		    if (SUCCEEDED(hr))
			hr = SetVariantFromSVEx(THIS_ *psv, pElement, cp, lcid);
		}
		else {
		    hr = SafeArrayPtrOfIndex(psa, pix, &V_BYREF(pElement));
		    if (SUCCEEDED(hr))
			hr = AssignVariantFromSV(THIS_ *psv, pElement,
						 cp, lcid);
		}
		if (hr == DISP_E_BADINDEX)
		    warn(MY_VERSION ": SetSafeArrayFromAV bad index");
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

    SafeArrayUnlock(psa);

    Safefree(pav);
    Safefree(pix);
    Safefree(plen);

    return hr;
}

HRESULT
SetVariantFromSVEx(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp, LCID lcid)
{
    HRESULT hr = S_OK;
    VariantClear(pVariant);

    /* XXX requirement to call mg_get() may change in Perl > 5.005 */
    if (SvGMAGICAL(sv))
	mg_get(sv);

    /* Objects */
    if (SvROK(sv)) {
	if (sv_derived_from(sv, szWINOLE)) {
	    WINOLEOBJECT *pObj = GetOleObject(THIS_ sv);
	    if (pObj) {
		pObj->pDispatch->AddRef();
		V_VT(pVariant) = VT_DISPATCH;
		V_DISPATCH(pVariant) = pObj->pDispatch;
		return S_OK;
	    }
	    return E_POINTER;
	}

	if (sv_derived_from(sv, szWINOLEVARIANT)) {
	    WINOLEVARIANTOBJECT *pVarObj =
		GetOleVariantObject(THIS_ sv);

	    if (pVarObj) {
		/* XXX Should we use VariantCopyInd? */
		hr = VariantCopy(pVariant, &pVarObj->variant);
	    }
	    else
		hr = E_POINTER;
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

	    if (psv && SvROK(*psv) && SvTYPE(SvRV(*psv)) == SVt_PVAV) {
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
	if (psa)
	    hr = SetSafeArrayFromAV(THIS_ (AV*)sv, VT_VARIANT, psa, dim,
				    cp, lcid);
	else
	    hr = E_OUTOFMEMORY;

	Safefree(pav);
	Safefree(pix);
	Safefree(plen);
	Safefree(psab);

	if (SUCCEEDED(hr)) {
	    V_VT(pVariant) = VT_VARIANT | VT_ARRAY;
	    V_ARRAY(pVariant) = psa;
	}
	else if (psa)
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
	V_VT(pVariant) = VT_BSTR;
	V_BSTR(pVariant) = AllocOleString(THIS_ SvPVX(sv), SvCUR(sv), cp);
    }
    else {
	V_VT(pVariant) = VT_ERROR;
	V_ERROR(pVariant) = DISP_E_PARAMNOTFOUND;
    }

    return hr;

}   /* SetVariantFromSVEx */

HRESULT
SetVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp)
{
    /* old API for PerlScript compatibility */
    return SetVariantFromSVEx(THIS_ sv, pVariant, cp, lcidDefault);
}   /* SetVariantFromSV */

HRESULT
AssignVariantFromSV(CPERLarg_ SV* sv, VARIANT *pVariant, UINT cp, LCID lcid)
{
    /* This function is similar to SetVariantFromSVEx except that
     * it does NOT choose the variant type itself.
     */
    HRESULT hr = S_OK;
    VARTYPE vt = V_VT(pVariant);
    /* sv must NOT be Nullsv unless vt is VT_EMPTY, VT_NULL or VT_DISPATCH */

#   define ASSIGN(vartype,perltype)                           \
        if (vt & VT_BYREF) {                                  \
            *V_##vartype##REF(pVariant) = Sv##perltype##(sv); \
        } else {                                              \
            V_##vartype(pVariant) = Sv##perltype##(sv);       \
        }

    /* XXX requirement to call mg_get() may change in Perl > 5.005 */
    if (sv && SvGMAGICAL(sv))
	mg_get(sv);

    if (vt & VT_ARRAY) {
	SAFEARRAY *psa;
	if (V_ISBYREF(pVariant))
	    psa = *V_ARRAYREF(pVariant);
	else
	    psa = V_ARRAY(pVariant);

	UINT cDims = SafeArrayGetDim(psa);
	if ((vt & VT_TYPEMASK) != VT_UI1 || cDims != 1 || !sv || !SvPOK(sv)) {
	    warn(MY_VERSION ": AssignVariantFromSV() cannot assign to "
		 "VT_ARRAY variant");
	    return E_INVALIDARG;
	}

	char *pDest;
	STRLEN len;
	char *pSrc = SvPV(sv, len);
	HRESULT hr = SafeArrayAccessData(psa, (void**)&pDest);
	if (SUCCEEDED(hr)) {
	    long lLower, lUpper;
	    SafeArrayGetLBound(psa, 1, &lLower);
	    SafeArrayGetUBound(psa, 1, &lUpper);

	    long lLength = 1 + lUpper-lLower;
	    len = (len < lLength ? len : lLength);
	    memcpy(pDest, pSrc, len);
	    if (lLength > len)
		memset(pDest+len, 0, lLength-len);

	    SafeArrayUnaccessData(psa);
	}
	return hr;
    }

    switch(vt & VT_TYPEMASK) {
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
	VARIANT variant;
	if (SvIOK(sv)) {
	    V_VT(&variant) = VT_I4;
	    V_I4(&variant) = SvIV(sv);
	}
	else if (SvNOK(sv)) {
	    V_VT(&variant) = VT_R8;
	    V_R8(&variant) = SvNV(sv);
	}
	else {
	    STRLEN len;
	    char *ptr = SvPV(sv, len);
	    V_VT(&variant) = VT_BSTR;
	    V_BSTR(&variant) = AllocOleString(THIS_ ptr, len, cp);
	}

	VARTYPE vt_base = vt & ~VT_BYREF;
	hr = VariantChangeTypeEx(&variant, &variant, lcid, 0, vt_base);
	if (SUCCEEDED(hr)) {
	    if (vt_base == VT_CY) {
		if (vt & VT_BYREF)
		    *V_CYREF(pVariant) = V_CY(&variant);
		else
		    V_CY(pVariant) = V_CY(&variant);
	    }
	    else {
		if (vt & VT_BYREF)
		    *V_DATEREF(pVariant) = V_DATE(&variant);
		else
		    V_DATE(pVariant) = V_DATE(&variant);
	    }
	}
	VariantClear(&variant);
	break;
    }

    case VT_BSTR:
    {
	STRLEN len;
	char *ptr = SvPV(sv, len);
	BSTR bstr = AllocOleString(THIS_ ptr, len, cp);

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
	if (vt & VT_BYREF) {
	    if (*V_DISPATCHREF(pVariant))
		(*V_DISPATCHREF(pVariant))->Release();
	    *V_DISPATCHREF(pVariant) = NULL;
	}
	else {
	    if (V_DISPATCH(pVariant))
		V_DISPATCH(pVariant)->Release();
	    V_DISPATCH(pVariant) = NULL;
	}
	if (sv_isobject(sv)) {
	    /* Argument MUST be a valid Perl OLE object! */
	    WINOLEOBJECT *pObj = GetOleObject(THIS_ sv);
	    if (pObj) {
		pObj->pDispatch->AddRef();
		if (vt & VT_BYREF)
		    *V_DISPATCHREF(pVariant) = pObj->pDispatch;
		else
		    V_DISPATCH(pVariant) = pObj->pDispatch;
	    }
	}
	break;

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
	if (vt & VT_BYREF)
	    hr = SetVariantFromSVEx(THIS_ sv, V_VARIANTREF(pVariant), cp, lcid);
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
	WINOLEOBJECT *pObj = GetOleObject(THIS_ sv);
	if (pObj) {
	    IUnknown *punk;
	    hr = pObj->pDispatch->QueryInterface(IID_IUnknown, (void**)&punk);
	    if (SUCCEEDED(hr)) {
		if (vt & VT_BYREF) {
		    if (*V_UNKNOWNREF(pVariant))
			(*V_UNKNOWNREF(pVariant))->Release();
		    *V_UNKNOWNREF(pVariant) = punk;
		}
		else {
		    if (V_UNKNOWN(pVariant))
			V_UNKNOWN(pVariant)->Release();
		    V_UNKNOWN(pVariant) = punk;
		}
	    }
	}
	break;
    }

    case VT_DECIMAL:
    {
	STRLEN len;
	char *ptr = SvPV(sv, len);

	VARIANT variant;
	VariantInit(&variant);
	V_VT(&variant) = VT_BSTR;
	V_BSTR(&variant) = AllocOleString(THIS_ ptr, len, cp);

	hr = VariantChangeTypeEx(&variant, &variant, lcid, 0, VT_DECIMAL);
	if (SUCCEEDED(hr)) {
	    if (vt & VT_BYREF)
		*V_DECIMALREF(pVariant) = V_DECIMAL(&variant);
	    else
		V_DECIMAL(pVariant) = V_DECIMAL(&variant);
	}
	VariantClear(&variant);
	break;
    }

    case VT_UI1:
	if (SvIOK(sv)) {
	    ASSIGN(UI1, IV);
	}
	else {
	    char *ptr = SvPV_nolen(sv);
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
	    ReportOleError(THIS_ stash, hr);
	}

	AddToObjectChain(THIS_ (OBJECTHEADER*)pVarObj, WINOLEVARIANT_MAGIC);
	SV *classname = newSVpv(HvNAME(stash), 0);
	sv_catpvn(classname, "::Variant", 9);
	sv_setref_pv(sv, SvPVX(classname), pVarObj);
	SvREFCNT_dec(classname);
	return hr;
    }

    while (vt == (VT_VARIANT|VT_BYREF)) {
	pVariant = V_VARIANTREF(pVariant);
	vt = V_VT(pVariant);
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
	if (dim == 1 && (vt & VT_TYPEMASK) == VT_UI1) {
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
	New(0, pav,         dim, AV*);

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

		SV *val = newSV(0);
		hr = SetSVFromVariantEx(THIS_ &variant, val, stash);
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
	UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

	if (V_ISBYREF(pVariant))
	    sv_setwide(THIS_ sv, *V_BSTRREF(pVariant), cp);
	else
	    sv_setwide(THIS_ sv, V_BSTR(pVariant), cp);
	break;
    }

    case VT_ERROR:
    case VT_DATE:
    {
	SV *classname;
	WINOLEVARIANTOBJECT *pVarObj;
	Newz(0, pVarObj, 1, WINOLEVARIANTOBJECT);
	VariantInit(&pVarObj->variant);
	VariantInit(&pVarObj->byref);
	hr = VariantCopy(&pVarObj->variant, pVariant);
	if (FAILED(hr)) {
	    Safefree(pVarObj);
	    ReportOleError(THIS_ stash, hr, NULL, NULL);
            break;
	}

	AddToObjectChain(THIS_ (OBJECTHEADER*)pVarObj, WINOLEVARIANT_MAGIC);
	classname = newSVpv(HvNAME(stash), 0);
	sv_catpvn(classname, "::Variant", 9);
	sv_setref_pv(sv, SvPVX(classname), pVarObj);
	SvREFCNT_dec(classname);
 	break;
    }

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

	if (pDispatch) {
	    pDispatch->AddRef();
	    sv_setsv(sv, CreatePerlObject(THIS_ stash, pDispatch, NULL));
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

	if (punk &&
	    SUCCEEDED(punk->QueryInterface(IID_IDispatch, (void**)&pDispatch)))
	{
	    sv_setsv(sv, CreatePerlObject(THIS_ stash, pDispatch, NULL));
	}
	break;
    }

    case VT_DECIMAL:
    {
	VARIANT variant;
	VariantInit(&variant);
	hr = VariantChangeTypeEx(&variant, pVariant, lcidDefault, 0, VT_R8);
	if (SUCCEEDED(hr) && V_VT(&variant) == VT_R8)
            sv_setnv(sv, V_R8(&variant));
	VariantClear(&variant);
	break;
    }

    case VT_CY:
    default:
    {
	LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
	UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);
	VARIANT variant;

	VariantInit(&variant);
	hr = VariantChangeTypeEx(&variant, pVariant, lcid, 0, VT_BSTR);
	if (SUCCEEDED(hr) && V_VT(&variant) == VT_BSTR)
	    sv_setwide(THIS_ sv, V_BSTR(&variant), cp);
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
    return SetSVFromVariantEx(THIS_ pVariant, sv, stash);
}

IV
GetLocaleNumber(CPERLarg_ HV *hv, char *key, LCID lcid, LCTYPE lctype)
{
    if (hv) {
	SV **psv = hv_fetch(hv, key, strlen(key), FALSE);
	if (psv)
	    return SvIV(*psv);
    }

    char *info;
    int len = GetLocaleInfo(lcid, lctype, NULL, 0);
    New(0, info, len, char);
    GetLocaleInfo(lcid, lctype, info, len);
    IV number = atol(info);
    Safefree(info);
    return number;
}

char *
GetLocaleString(CPERLarg_ HV *hv, char *key, LCID lcid, LCTYPE lctype)
{
    if (hv) {
	SV **psv = hv_fetch(hv, key, strlen(key), FALSE);
	if (psv)
	    return SvPV_nolen(*psv);
    }

    int len = GetLocaleInfo(lcid, lctype, NULL, 0);
    SV *sv = sv_2mortal(newSV(len));
    GetLocaleInfo(lcid, lctype, SvPVX(sv), len);
    return SvPVX(sv);
}

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

	DBG(("Initialize dwCoInit=%d\n", dwCoInit));

	if (dwCoInit == COINIT_OLEINITIALIZE) {
	    hr = OleInitialize(NULL);
	    if (SUCCEEDED(hr))
		g_pfnCoUninitialize = &OleUninitialize;
	}
	else if (dwCoInit != COINIT_NO_INITIALIZE) {
	    if (g_pfnCoInitializeEx)
		hr = g_pfnCoInitializeEx(NULL, dwCoInit);
	    else
		hr = CoInitialize(NULL);

	    if (SUCCEEDED(hr))
		g_pfnCoUninitialize = &CoUninitialize;
	}

	if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
	    ReportOleError(THIS_ stash, hr);
    }

    LeaveCriticalSection(&g_CriticalSection);

}   /* Initialize */

void
Uninitialize(CPERLarg_ PERINTERP *pInterp)
{
    DBG(("Uninitialize\n"));
    EnterCriticalSection(&g_CriticalSection);
    if (g_bInitialized) {
	OBJECTHEADER *pHeader = g_pObj;
	while (pHeader) {
	    DBG(("Zombiefy object |%lx| lMagic=%lx\n",
		 pHeader, pHeader->lMagic));

	    switch (pHeader->lMagic) {
	    case WINOLE_MAGIC:
		ReleasePerlObject(THIS_ (WINOLEOBJECT*)pHeader);
		break;

	    case WINOLEENUM_MAGIC: {
		WINOLEENUMOBJECT *pEnumObj = (WINOLEENUMOBJECT*)pHeader;
		if (pEnumObj->pEnum) {
		    pEnumObj->pEnum->Release();
		    pEnumObj->pEnum = NULL;
		}
		break;
	    }

	    case WINOLEVARIANT_MAGIC: {
		WINOLEVARIANTOBJECT *pVarObj = (WINOLEVARIANTOBJECT*)pHeader;
		VariantClear(&pVarObj->byref);
		VariantClear(&pVarObj->variant);
		break;
	    }

	    case WINOLETYPELIB_MAGIC: {
		WINOLETYPELIBOBJECT *pObj = (WINOLETYPELIBOBJECT*)pHeader;
		if (pObj->pTypeLib) {
		    pObj->pTypeLib->Release();
		    pObj->pTypeLib = NULL;
		}
		break;
	    }

	    case WINOLETYPEINFO_MAGIC: {
		WINOLETYPEINFOOBJECT *pObj = (WINOLETYPEINFOOBJECT*)pHeader;
		if (pObj->pTypeInfo) {
		    pObj->pTypeInfo->Release();
		    pObj->pTypeInfo = NULL;
		}
		break;
	    }

	    default:
		DBG(("Unknown magic number: %08lx", pHeader->lMagic));
		break;
	    }
	    pHeader = pHeader->pNext;
	}

	DBG(("CoUninitialize\n"));
	if (g_pfnCoUninitialize)
	    g_pfnCoUninitialize();
	g_bInitialized = FALSE;
    }
    LeaveCriticalSection(&g_CriticalSection);

}   /* Uninitialize */

static void
AtExit(pTHX_ CPERLarg_ void *pVoid)
{
    PERINTERP *pInterp = (PERINTERP*)pVoid;

    DeleteCriticalSection(&g_CriticalSection);
    if (g_hOLE32)
	FreeLibrary(g_hOLE32);
    if (g_hHHCTRL)
	FreeLibrary(g_hHHCTRL);
#if defined(MULTIPLICITY) || defined(PERL_OBJECT)
    Safefree(pInterp);
#endif
    DBG(("AtExit done\n"));

}   /* AtExit */

void
Bootstrap(CPERLarg)
{
    dSP;
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
    if (g_hOLE32) {
	g_pfnCoInitializeEx = (FNCOINITIALIZEEX*)
	    GetProcAddress(g_hOLE32, "CoInitializeEx");
	g_pfnCoCreateInstanceEx = (FNCOCREATEINSTANCEEX*)
	    GetProcAddress(g_hOLE32, "CoCreateInstanceEx");
    }

    g_hHHCTRL = NULL;
    g_pfnHtmlHelp = NULL;

    SV *cmd = newSVpv("", 0);
    sv_setpvf(cmd, "END { %s->Uninitialize(%d); }", szWINOLE, WINOLE_MAGIC );

    PUSHMARK(sp);
    perl_eval_sv(cmd, G_DISCARD);
    SPAGAIN;

    SvREFCNT_dec(cmd);

#if (PATCHLEVEL > 4) || (SUBVERSION >= 68)
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
    MessageLoop = 3
    QuitMessageLoop = 4
    FreeUnusedLibraries = 5
PPCODE:
{
    char *paszMethod[] = {"Initialize", "Uninitialize", "SpinMessageLoop",
                          "MessageLoop", "QuitMessageLoop",
			  "FreeUnusedLibraries"};

    if (CallObjectMethod(THIS_ mark, ax, items, paszMethod[ix]))
	return;

    DBG(("Win32::OLE->%s()\n", paszMethod[ix]));

    if (items == 0) {
        warn("Win32::OLE->%s must be called as class method", paszMethod[ix]);
	XSRETURN_EMPTY;
    }

    HV *stash = gv_stashsv(ST(0), TRUE);
    SetLastOleError(THIS_ stash);

    switch (ix) {
    case 0: {		// Initialize
	DWORD dwCoInit = COINIT_MULTITHREADED;
	if (items > 1 && SvOK(ST(1)))
	    dwCoInit = SvIV(ST(1));

	Initialize(THIS_ gv_stashsv(ST(0), TRUE), dwCoInit);
	break;
    }
    case 1: {		// Uninitialize
	dPERINTERP;
	Uninitialize(THIS_ INTERP);
	break;
    }
    case 2:		// SpinMessageLoop
	SpinMessageLoop();
	break;

    case 3: {		// MessageLoop
	MSG msg;
	DBG(("MessageLoop\n"));
	while (GetMessage(&msg, NULL, 0, 0)) {
	    if (msg.hwnd == NULL && msg.message == WM_USER)
		break;
	    TranslateMessage(&msg);
	    DispatchMessage(&msg);
	}
	break;
    }
    case 4:		// QuitMessageLoop
	PostThreadMessage(GetCurrentThreadId(), WM_USER, 0, 0);
	break;

    case 5:		// FreeUnusedLibraries
	CoFreeUnusedLibraries();
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

    if (CallObjectMethod(THIS_ mark, ax, items, "new"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->new(PROGID[,DESTROY])");
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(THIS_ stash);
    SetLastOleError(THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(THIS_ ST(2), "Win32::OLE->new");

    ST(0) = &PL_sv_undef;

    /* normal case: no DCOM */
    char *pszProgID;
    if (!SvROK(progid) || SvTYPE(SvRV(progid)) != SVt_PVAV) {
	pszProgID = SvPV_nolen(progid);
	pBuffer = GetWideChar(THIS_ pszProgID, Buffer, OLE_BUF_SIZ, cp);
	if (isalpha(pszProgID[0]))
	    hr = CLSIDFromProgID(pBuffer, &clsid);
	else
	    hr = CLSIDFromString(pBuffer, &clsid);
	ReleaseBuffer(THIS_ pBuffer, Buffer);
	if (SUCCEEDED(hr)) {
	    hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER,
				  IID_IDispatch, (void**)&pDispatch);
	}
	if (!CheckOleError(THIS_ stash, hr)) {
	    ST(0) = CreatePerlObject(THIS_ stash, pDispatch, destroy);
	    DBG(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
	}
	XSRETURN(1);
    }

    /* DCOM might not exist on Win95 (and does not on NT 3.5) */
    dPERINTERP;
    if (!g_pfnCoCreateInstanceEx) {
	hr = HRESULT_FROM_WIN32(ERROR_SERVICE_DOES_NOT_EXIST);
	ReportOleError(THIS_ stash, hr);
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
	pszHost = SvPVX(host);
	if (IsLocalMachine(THIS_ pszHost))
	    pszHost = NULL;
    }

    /* determine CLSID */
    pszProgID = SvPV_nolen(progid);
    pBuffer = GetWideChar(THIS_ pszProgID, Buffer, OLE_BUF_SIZ, cp);
    if (isalpha(pszProgID[0])) {
	hr = CLSIDFromProgID(pBuffer, &clsid);
	if (FAILED(hr) && pszHost)
	    hr = CLSIDFromRemoteRegistry(THIS_ pszHost, pszProgID, &clsid);
    }
    else
        hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(THIS_ pBuffer, Buffer);
    if (FAILED(hr)) {
	ReportOleError(THIS_ stash, hr);
	XSRETURN(1);
    }

    /* setup COSERVERINFO & MULTI_QI parameters */
    DWORD clsctx = CLSCTX_REMOTE_SERVER;
    COSERVERINFO ServerInfo;
    OLECHAR ServerName[OLE_BUF_SIZ];
    MULTI_QI multi_qi;

    Zero(&ServerInfo, 1, COSERVERINFO);
    if (pszHost)
	ServerInfo.pwszName = GetWideChar(THIS_ pszHost, ServerName,
					  OLE_BUF_SIZ, cp);
    else
	clsctx = CLSCTX_SERVER;

    Zero(&multi_qi, 1, MULTI_QI);
    multi_qi.pIID = &IID_IDispatch;

    /* create instance on remote server */
    hr = g_pfnCoCreateInstanceEx(clsid, NULL, clsctx, &ServerInfo,
				  1, &multi_qi);
    ReleaseBuffer(THIS_ ServerInfo.pwszName, ServerName);
    if (!CheckOleError(THIS_ stash, hr)) {
	pDispatch = (IDispatch*)multi_qi.pItf;
	ST(0) = CreatePerlObject(THIS_ stash, pDispatch, destroy);
	DBG(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
    }
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(THIS_ self, TRUE);
    DBG(("Win32::OLE::DESTROY |%lx| |%lx|\n", pObj,
	 pObj ? pObj->pDispatch : NULL));
    if (pObj) {
	ReleasePerlObject(THIS_ pObj);
	pObj->bDestroyed = TRUE;
    }
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

    pObj = GetOleObject(THIS_ self);
    if (!pObj) {
	XSRETURN(1);
    }

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(THIS_ stash);

    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

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
	    hr = GetHashedDispID(THIS_ pObj, buffer, length, dispID, lcid, cp);
	    if (FAILED(hr)) {
		if (PL_hints & HINT_STRICT_SUBS) {
		    err = newSVpvf(" in GetIDsOfNames of \"%s\"", buffer);
		    ReportOleError(THIS_ stash, hr, NULL, sv_2mortal(err));
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

	New(0, rghe, dispParams.cNamedArgs, HE*);
	New(0, dispParams.rgdispidNamedArgs, dispParams.cNamedArgs, DISPID);
	New(0, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	for (index = 0 ; index < dispParams.cArgs ; ++index)
	    VariantInit(&dispParams.rgvarg[index]);

	New(0, rgszNames, 1+dispParams.cNamedArgs, OLECHAR*);
	New(0, rgdispids, 1+dispParams.cNamedArgs, DISPID);

	rgszNames[0] = AllocOleString(THIS_ buffer, length, cp);
	hv_iterinit(hv);
	for (index = 0; index < dispParams.cNamedArgs; ++index) {
	    rghe[index] = hv_iternext(hv);
	    char *pszName = hv_iterkey(rghe[index], &len);
	    rgszNames[1+index] = AllocOleString(THIS_ pszName, len, cp);
	}

	hr = pObj->pDispatch->GetIDsOfNames(IID_NULL, rgszNames,
			      1+dispParams.cNamedArgs, lcid, rgdispids);

	if (SUCCEEDED(hr)) {
	    for (index = 0; index < dispParams.cNamedArgs; ++index) {
		dispParams.rgdispidNamedArgs[index] = rgdispids[index+1];
		hr = SetVariantFromSVEx(THIS_ hv_iterval(hv, rghe[index]),
					&dispParams.rgvarg[index], cp, lcid);
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
	if (!dispParams.rgvarg) {
	    New(0, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	    for (index = 0 ; index < dispParams.cArgs ; ++index)
		VariantInit(&dispParams.rgvarg[index]);
	}

	for(index = dispParams.cNamedArgs; index < dispParams.cArgs; ++index) {
	    SV *sv = ST(items-1-(index-dispParams.cNamedArgs));
	    hr = SetVariantFromSVEx(THIS_ sv, &dispParams.rgvarg[index],
				    cp, lcid);
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
		GetOleVariantObject(THIS_ retval);

	    if (pVarObj) {
		VariantClear(&pVarObj->byref);
		VariantClear(&pVarObj->variant);
		VariantCopy(&pVarObj->variant, &result);
		ST(0) = &PL_sv_yes;
	    }
	}
	else {
	    hr = SetSVFromVariantEx(THIS_ &result, retval, stash);
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
		sv_catpvf(err, " argument %d", dispParams.cArgs - argErr);
	}
    }

 Cleanup:
    VariantClear(&result);
    if (dispParams.cArgs != 0 && dispParams.rgvarg) {
	for(index = 0; index < dispParams.cArgs; ++index)
	    VariantClear(&dispParams.rgvarg[index]);
	Safefree(dispParams.rgvarg);
    }
    Safefree(rghe);
    if (dispParams.rgdispidNamedArgs != &dispIDParam)
	Safefree(dispParams.rgdispidNamedArgs);

    CheckOleError(THIS_ stash, hr, &excepinfo, err);

    XSRETURN(1);
}

void
EnumAllObjects(...)
PPCODE:
{
    if (CallObjectMethod(THIS_ mark, ax, items, "EnumAllObjects"))
	return;

    if (items > 2) {
	warn("Usage: Win32::OLE->EnumAllObjects([CALLBACK])");
	XSRETURN_EMPTY;
    }

    if (items == 2 && (!SvROK(ST(1)) || SvTYPE(SvRV(ST(1))) != SVt_PVCV)) {
	warn(MY_VERSION "Win32::OLE->EnumAllObjects: "
	     "CALLBACK must be a CODE ref");
	XSRETURN_EMPTY;
    }

    dPERINTERP;
    IV count = 0;
    OBJECTHEADER *pHeader = g_pObj;
    SV *callback = (items == 2) ? ST(1) : NULL;

    while (pHeader) {
	if (pHeader->lMagic == WINOLE_MAGIC) {
	    ++count;
	    if (callback) {
		WINOLEOBJECT *pObj = (WINOLEOBJECT*)pHeader;;
		SV *self = newRV_inc((SV*)pObj->self);
		if (Gv_AMG(SvSTASH(pObj->self)))
		    SvAMAGIC_on(self);

		ENTER;
		SAVETMPS;
		PUSHMARK(sp);
		XPUSHs(sv_2mortal(self));
		PUTBACK;
		perl_call_sv(callback, G_DISCARD);
		SPAGAIN;
		FREETMPS;
		LEAVE;
	    }
	}
	pHeader = pHeader->pNext;
    }
    XSRETURN_IV(count);
}

void
Forward(self,method)
    SV *self
    SV *method
PPCODE:
{
    if (CallObjectMethod(THIS_ mark, ax, items, "Forward"))
	return;

    if (!SvROK(method) || SvTYPE(SvRV(method)) != SVt_PVCV) {
	warn("Win32::OLE->Forward: method must be a CODE ref");
	XSRETURN_EMPTY;
    }

    HV *stash = gv_stashsv(self, TRUE);
    IDispatch *pDispatch = new Forwarder(THIS_ stash, method);
    ST(0) = CreatePerlObject(THIS_ stash, pDispatch, NULL);
    XSRETURN(1);
}

void
GetActiveObject(...)
PPCODE:
{
    CLSID clsid;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    char *buffer;
    HRESULT hr;
    IUnknown *pUnknown;
    IDispatch *pDispatch;

    if (CallObjectMethod(THIS_ mark, ax, items, "GetActiveObject"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->GetActiveObject(PROGID[,DESTROY])");
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *progid = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(THIS_ stash);
    SetLastOleError(THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(THIS_ ST(2),
				       "Win32::OLE->GetActiveObject");

    buffer = SvPV_nolen(progid);
    pBuffer = GetWideChar(THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    if (isalpha(buffer[0]))
        hr = CLSIDFromProgID(pBuffer, &clsid);
    else
        hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(THIS_ pBuffer, Buffer);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = GetActiveObject(clsid, 0, &pUnknown);
    /* Don't call CheckOleError! Return "undef" for "Server not running" */
    if (FAILED(hr))
	XSRETURN_EMPTY;

    hr = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(THIS_ stash, pDispatch, destroy);
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

    if (CallObjectMethod(THIS_ mark, ax, items, "GetObject"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->GetObject(PATHNAME[,DESTROY])");
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *pathname = ST(1);
    SV *destroy = NULL;
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    Initialize(THIS_ stash);
    SetLastOleError(THIS_ stash);

    if (items == 3)
	destroy = CheckDestroyFunction(THIS_ ST(2), "Win32::OLE->GetObject");

    hr = CreateBindCtx(0, &pBindCtx);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    buffer = SvPV_nolen(pathname);
    pBuffer = GetWideChar(THIS_ buffer, Buffer, OLE_BUF_SIZ, cp);
    hr = MkParseDisplayName(pBindCtx, pBuffer, &ulEaten, &pMoniker);
    ReleaseBuffer(THIS_ pBuffer, Buffer);
    if (FAILED(hr)) {
	pBindCtx->Release();
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "after character %lu in \"%s\"", ulEaten, buffer);
	ReportOleError(THIS_ stash, hr, NULL, sv);
	XSRETURN_EMPTY;
    }

    hr = pMoniker->BindToObject(pBindCtx, NULL, IID_IDispatch,
				 (void**)&pDispatch);
    pBindCtx->Release();
    pMoniker->Release();
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(THIS_ stash, pDispatch, destroy);
    XSRETURN(1);
}

void
GetTypeInfo(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    HV *stash = gv_stashsv(self, TRUE);
    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);

    SetLastOleError(THIS_ stash);
    HRESULT hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ stash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr));
    XSRETURN(1);
}

void
QueryInterface(self,itf)
    SV *self
    SV *itf
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    IID iid;
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;

    // XXX support GUIDs in addition to names too
    char *pszItf = SvPV_nolen(itf);

    DBG(("QueryInterface(%s)\n", pszItf));
    HV *stash = SvSTASH(pObj->self);
    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    SetLastOleError(THIS_ stash);

    HRESULT hr = FindIID(THIS_ pObj, pszItf, &iid, NULL, cp, lcid);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    IUnknown *pUnknown;
    hr = pObj->pDispatch->QueryInterface(iid, (void**)&pUnknown);
    DBG(("  QueryInterface(iid): 0x%08x\n", hr));
    if (CheckOleError(THIS_ stash, hr))
        XSRETURN_EMPTY;

    IDispatch *pDispatch;
    hr = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    DBG(("  QueryInterface(IDispatch): 0x%08x\n", hr));
    pUnknown->Release();
    if (CheckOleError(THIS_ stash, hr))
        XSRETURN_EMPTY;

    ST(0) = CreatePerlObject(THIS_ stash, pDispatch, NULL);
    DBG(("Win32::OLE::QueryInterface |%lx| |%lx|\n", ST(0), pDispatch));
    XSRETURN(1);
}

void
QueryObjectType(...)
PPCODE:
{
    if (CallObjectMethod(THIS_ mark, ax, items, "QueryObjectType"))
	return;

    if (items != 2) {
	warn("Usage: Win32::OLE->QueryObjectType(OBJECT)");
	XSRETURN_EMPTY;
    }

    SV *object = ST(1);

    if (!sv_isobject(object) || !sv_derived_from(object, szWINOLE)) {
	warn("Win32::OLE->QueryObjectType: object is not a Win32::OLE object");
	XSRETURN_EMPTY;
    }

    WINOLEOBJECT *pObj = GetOleObject(THIS_ object);
    if (!pObj)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int count;
    BSTR bstr;

    HRESULT hr = pObj->pDispatch->GetTypeInfoCount(&count);
    if (FAILED(hr) || count == 0)
	XSRETURN_EMPTY;

    HV *stash = gv_stashsv(ST(0), TRUE);
    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    SetLastOleError(THIS_ stash);
    hr = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    /* Return ('TypeLib Name', 'Class Name') in array context */
    if (GIMME_V == G_ARRAY) {
	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(THIS_ stash, hr);
	    XSRETURN_EMPTY;
	}

	hr = pTypeLib->GetDocumentation(-1, &bstr, NULL, NULL, NULL);
	pTypeLib->Release();
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(THIS_ stash, hr);
	    XSRETURN_EMPTY;
	}

	PUSHs(sv_2mortal(sv_setwide(THIS_ NULL, bstr, cp)));
	SysFreeString(bstr);
    }

    hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstr, NULL, NULL, NULL);
    pTypeInfo->Release();
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    PUSHs(sv_2mortal(sv_setwide(THIS_ NULL, bstr, cp)));
    SysFreeString(bstr);
}

void
WithEvents(...)
PPCODE:
{
    if (CallObjectMethod(THIS_ mark, ax, items, "WithEvents"))
	return;

    if (items < 2) {
	warn("Usage: Win32::OLE->WithEvents(OBJECT [, HANDLER [, INTERFACE]])");
	XSRETURN_EMPTY;
    }

    WINOLEOBJECT *pObj = GetOleObject(THIS_ ST(1));
    if (!pObj)
	XSRETURN_EMPTY;

    // disconnect previous event handler
    if (pObj->pEventSink) {
	pObj->pEventSink->Unadvise();
	pObj->pEventSink = NULL;
    }

    if (items == 2)
	XSRETURN_EMPTY;

    SV *handler = ST(2);
    HV *stash = SvSTASH(pObj->self);

    // make sure we are running in a single threaded apartment
    HRESULT hr = CoInitialize(NULL);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;
    CoUninitialize();

    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);
    SetLastOleError(THIS_ stash);

    IID iid;
    ITypeInfo *pTypeInfo = NULL;

    // Interfacename specified?
    if (items > 3) {
	SV *itf = ST(3);
	if (sv_isobject(itf) && sv_derived_from(itf, szWINOLETYPEINFO)) {
	    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ itf);
	    if (!pObj)
		XSRETURN_EMPTY;

	    if (pObj->pTypeAttr->typekind == TKIND_DISPATCH) {
		iid = (IID)pObj->pTypeAttr->guid;
		pTypeInfo = pObj->pTypeInfo;
		pTypeInfo->AddRef();
	    }
	    else if (pObj->pTypeAttr->typekind == TKIND_COCLASS) {
		// Enumerate all implemented types of the COCLASS
		for (UINT i=0; i < pObj->pTypeAttr->cImplTypes; i++) {
		    int iFlags;
		    hr = pObj->pTypeInfo->GetImplTypeFlags(i, &iFlags);
		    DBG(("GetImplTypeFlags: hr=0x%08x i=%d iFlags=%d\n", hr, i, iFlags));
		    if (FAILED(hr))
			continue;

		    // looking for the [default] [source]
		    // we just hope that it is a dispinterface :-)
		    if ((iFlags & IMPLTYPEFLAG_FDEFAULT) &&
			(iFlags & IMPLTYPEFLAG_FSOURCE))
		    {
			HREFTYPE hRefType = NULL;
			hr = pObj->pTypeInfo->GetRefTypeOfImplType(i, &hRefType);
			DBG(("GetRefTypeOfImplType: hr=0x%08x\n", hr));
			if (FAILED(hr))
			    continue;
			hr = pObj->pTypeInfo->GetRefTypeInfo(hRefType, &pTypeInfo);
			DBG(("GetRefTypeInfo: hr=0x%08x\n", hr));
			if (SUCCEEDED(hr))
			    break;
		    }
		}

		// Now that would be a bad surprise, if we didn't find it, wouldn't it?
		if (!pTypeInfo) {
		    if (SUCCEEDED(hr))
			hr = E_UNEXPECTED;
		}
		else {
		    // Determine IID of default source interface
		    TYPEATTR *pTypeAttr;
		    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		    if (SUCCEEDED(hr)) {
			iid = pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		    }
		    else
			pTypeInfo->Release();
		}
	    }
	    else {
		XSRETURN_EMPTY; /* set hr instead XXX error message */
	    }
	}
	else { /* interface _not_ a Win32::OLE::TypeInfo object */
	    char *pszItf = SvPV_nolen(itf);
	    if (isalpha(pszItf[0]))
		hr = FindIID(THIS_ pObj, pszItf, &iid, &pTypeInfo, cp, lcid);
	    else {
		OLECHAR Buffer[OLE_BUF_SIZ];
		OLECHAR *pBuffer = GetWideChar(THIS_ pszItf, Buffer, OLE_BUF_SIZ, cp);
		hr = IIDFromString(pBuffer, &iid);
		ReleaseBuffer(THIS_ pBuffer, Buffer);
	    }
	}
    }
    else
	hr = FindDefaultSource(THIS_ pObj, &iid, &pTypeInfo, cp, lcid);

    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    // Get IConnectionPointContainer interface
    IConnectionPointContainer *pContainer;
    hr = pObj->pDispatch->QueryInterface(IID_IConnectionPointContainer,
					 (void**)&pContainer);
    DBG(("QueryInterFace(IConnectionPointContainer): hr=0x%08x\n", hr));
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ stash, hr);
        XSRETURN_EMPTY;
    }

    // Find default source connection point
    IConnectionPoint *pConnectionPoint;
    hr = pContainer->FindConnectionPoint(iid, &pConnectionPoint);
    pContainer->Release();
    DBG(("FindConnectionPoint: hr=0x%08x\n", hr));
    if (FAILED(hr)) {
	if (pTypeInfo)
	    pTypeInfo->Release();
	ReportOleError(THIS_ stash, hr);
        XSRETURN_EMPTY;
    }

    // Connect our EventSink object to it
    pObj->pEventSink = new EventSink(THIS_ pObj, handler, iid, pTypeInfo);
    hr = pObj->pEventSink->Advise(pConnectionPoint);
    pConnectionPoint->Release();
    DBG(("Advise: hr=0x%08x\n", hr));
    if (FAILED(hr)) {
	if (pTypeInfo)
	    pTypeInfo->Release();
	pObj->pEventSink->Release();
	pObj->pEventSink = NULL;
	ReportOleError(THIS_ stash, hr);
    }

 #ifdef _DEBUG
    // Get IOleControl interface
    IOleControl *pOleControl;
    hr = pObj->pDispatch->QueryInterface(IID_IOleControl, (void**)&pOleControl);
    DBG(("QueryInterface(IOleControl): 0x%08x\n", hr));
    if (SUCCEEDED(hr)) {
	pOleControl->FreezeEvents(TRUE);
	pOleControl->FreezeEvents(FALSE);
	pOleControl->Release();
    }
 #endif

    XSRETURN_EMPTY;
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Tie

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(THIS_ self, TRUE);
    DBG(("Win32::OLE::Tie::DESTROY |%lx| |%lx|\n", pObj,
	 pObj ? pObj->pDispatch : NULL));

    if (pObj) {
	/* objects may be destroyed in the wrong order during global cleanup */
	if (!pObj->bDestroyed) {
	    DBG(("Win32::OLE::Tie::DESTROY: OLE object not yet destroyed\n"));
	    if (pObj->pDispatch) {
		/* make sure the reference to the tied hash is still valid */
		sv_unmagic((SV*)pObj->self, 'P');
		sv_magic((SV*)pObj->self, self, 'P', Nullch, 0);
		ReleasePerlObject(THIS_ pObj);
	    }
	    /* untie hash because we free the object *right now* */
	    sv_unmagic((SV*)pObj->self, 'P');
	}
	RemoveFromObjectChain(THIS_ (OBJECTHEADER*)pObj);
	Safefree(pObj);
    }
    DBG(("End of Win32::OLE::Tie::DESTROY\n"));
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
    STRLEN length;
    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    VARIANTARG propName;
    DISPID dispID = DISPID_VALUE;
    HRESULT hr;

    buffer = SvPV(key, length);
    if (strEQ(buffer, PERL_OLE_ID)) {
	DBG(("Win32::OLE::Tie::Fetch(0x%08x,'%s')\n", self, buffer));
	ST(0) = *hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
	XSRETURN(1);
    }

    WINOLEOBJECT *pObj = GetOleObject(THIS_ self);
    DBG(("Win32::OLE::Tie::Fetch(0x%08x,'%s',%d)\n", pObj, buffer));
    if (!pObj)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(THIS_ stash);

    ST(0) = &PL_sv_undef;
    VariantInit(&result);
    VariantInit(&propName);

    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    dispParams.cArgs = 0;
    dispParams.rgvarg = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    hr = GetHashedDispID(THIS_ pObj, buffer, length, dispID, lcid, cp);
    if (FAILED(hr)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(THIS_ stash, hr, NULL, sv_2mortal(err));
	    XSRETURN(1);
	}

	/* default method call: $self->{Key} ---> $self->Item('Key') */
	V_VT(&propName) = VT_BSTR;
	V_BSTR(&propName) = AllocOleString(THIS_ buffer, length, cp);
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
	ReportOleError(THIS_ stash, hr, &excepinfo, sv);
    }
    else {
	ST(0) = sv_newmortal();
	hr = SetSVFromVariantEx(THIS_ &result, ST(0), stash);
	VariantClear(&result);
	CheckOleError(THIS_ stash, hr);
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
    unsigned int argErr;
    STRLEN length;
    char *buffer;
    int index;
    HRESULT hr;
    EXCEPINFO excepinfo;
    DISPID dispID = DISPID_VALUE;
    DISPID dispIDParam = DISPID_PROPERTYPUT;
    DISPPARAMS dispParams;
    VARIANTARG propertyValue[2];
    SV *err = NULL;

    WINOLEOBJECT *pObj = GetOleObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(THIS_ stash);

    LCID lcid = QueryPkgVar(THIS_ stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

    dispParams.rgdispidNamedArgs = &dispIDParam;
    dispParams.rgvarg = propertyValue;
    dispParams.cNamedArgs = 1;
    dispParams.cArgs = 1;

    VariantInit(&propertyValue[0]);
    VariantInit(&propertyValue[1]);
    Zero(&excepinfo, 1, EXCEPINFO);

    buffer = SvPV(key, length);
    hr = GetHashedDispID(THIS_ pObj, buffer, length, dispID, lcid, cp);
    if (FAILED(hr)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(THIS_ stash, hr, NULL, sv_2mortal(err));
	    XSRETURN_EMPTY;
	}

	dispParams.cArgs = 2;
	V_VT(&propertyValue[1]) = VT_BSTR;
	V_BSTR(&propertyValue[1]) = AllocOleString(THIS_ buffer, length, cp);
    }

    hr = SetVariantFromSVEx(THIS_ value, &propertyValue[0], cp, lcid);
    if (SUCCEEDED(hr)) {
	USHORT wFlags = DISPATCH_PROPERTYPUT;

	/* objects are passed by reference */
	VARTYPE vt = V_VT(&propertyValue[0]) & VT_TYPEMASK;
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

    if (CheckOleError(THIS_ stash, hr, &excepinfo, err))
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
    WINOLEOBJECT *pObj = GetOleObject(THIS_ self);
    char *paszMethod[] = {"FIRSTKEY", "NEXTKEY", "FIRSTENUM", "NEXTENUM"};

    DBG(("%s called, pObj=%p\n", paszMethod[ix], pObj));
    if (!pObj)
	XSRETURN_EMPTY;

    HV *stash = SvSTASH(pObj->self);
    SetLastOleError(THIS_ stash);

    switch (ix) {
    case 0: /* FIRSTKEY */
	FetchTypeInfo(THIS_ pObj);
	pObj->PropIndex = 0;
    case 1: /* NEXTKEY */
	ST(0) = NextPropertyName(THIS_ pObj);
	break;

    case 2: /* FIRSTENUM */
	if (pObj->pEnum)
	    pObj->pEnum->Release();
	pObj->pEnum = CreateEnumVARIANT(THIS_ pObj);
    case 3: /* NEXTENUM */
	ST(0) = NextEnumElement(THIS_ pObj->pEnum, stash);
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
_LoadRegTypeLib(classid,major,minor,locale,typelib,codepage)
    SV *classid
    IV major
    IV minor
    SV *locale
    SV *typelib
    SV *codepage
PPCODE:
{
    ITypeLib *pTypeLib;
    TLIBATTR *pTLibAttr;
    CLSID clsid;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    HRESULT hr;
    LCID lcid = SvIOK(locale) ? SvIV(locale) : lcidDefault;
    UINT cp = SvIOK(codepage) ? SvIV(codepage) : cpDefault;
    HV *stash = gv_stashpv(szWINOLE, TRUE);
    unsigned int count;

    Initialize(THIS_ stash);
    SetLastOleError(THIS_ stash);

    char *pszBuffer = SvPV_nolen(classid);
    pBuffer = GetWideChar(THIS_ pszBuffer, Buffer, OLE_BUF_SIZ, cp);
    hr = CLSIDFromString(pBuffer, &clsid);
    ReleaseBuffer(THIS_ pBuffer, Buffer);
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = LoadRegTypeLib(clsid, major, minor, lcid, &pTypeLib);
    if (FAILED(hr) && SvPOK(typelib)) {
	/* typelib not registerd, try to read from file "typelib" */
	pszBuffer = SvPV_nolen(typelib);
	pBuffer = GetWideChar(THIS_ pszBuffer, Buffer, OLE_BUF_SIZ, cp);
	hr = LoadTypeLibEx(pBuffer, REGKIND_NONE, &pTypeLib);
	ReleaseBuffer(THIS_ pBuffer, Buffer);
    }
    if (CheckOleError(THIS_ stash, hr))
	XSRETURN_EMPTY;

    hr = pTypeLib->GetLibAttr(&pTLibAttr);
    if (FAILED(hr)) {
	pTypeLib->Release();
	ReportOleError(THIS_ stash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeLibObject(THIS_ pTypeLib, pTLibAttr));
    XSRETURN(1);
}

void
_Constants(typelib,caller)
    SV *typelib
    SV *caller
PPCODE:
{
    HRESULT hr;
    UINT cp = cpDefault;
    HV *stash = gv_stashpv(szWINOLE, TRUE);
    HV *hv;
    unsigned int count;

    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ typelib);
    if (!pObj)
	XSRETURN_EMPTY;

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
    count = pObj->pTypeLib->GetTypeInfoCount();
    for (int index=0 ; index < count ; ++index) {
	ITypeInfo *pTypeInfo;
	TYPEATTR  *pTypeAttr;

	hr = pObj->pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (CheckOleError(THIS_ stash, hr))
	    continue;

	hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(THIS_ stash, hr);
	    continue;
	}

	for (int iVar=0 ; iVar < pTypeAttr->cVars ; ++iVar) {
	    VARDESC *pVarDesc;

	    hr = pTypeInfo->GetVarDesc(iVar, &pVarDesc);
	    /* XXX LEAK alert */
	    if (CheckOleError(THIS_ stash, hr))
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
		if (CheckOleError(THIS_ stash, hr) || cName == 0 || !bstr)
		    continue;

		char *pszName = GetMultiByte(THIS_ bstr,
					     szName, sizeof(szName), cp);
		SV *sv = newSV(0);
		/* XXX LEAK alert */
		hr = SetSVFromVariantEx(THIS_ pVarDesc->lpvarValue, sv, stash);
		if (!CheckOleError(THIS_ stash, hr)) {
		    if (SvOK(caller)) {
			/* XXX check for valid symbol name */
			newCONSTSUB(hv, pszName, sv);
		    }
		    else
		        hv_store(hv, pszName, strlen(pszName), sv, 0);
		}
		SysFreeString(bstr);
		ReleaseBuffer(THIS_ pszName, szName);
	    }
	    pTypeInfo->ReleaseVarDesc(pVarDesc);
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeInfo->Release();
    }
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
		err = RegOpenKeyEx(hKeyVersion, szLangid, 0, KEY_READ,
				   &hKeyLangid);
		if (err != ERROR_SUCCESS)
		    continue;

		// Retrieve filename of type library
		char szFile[MAX_PATH+1];
		LONG cbFile = sizeof(szFile);
		err = RegQueryValue(hKeyLangid, "win32", szFile, &cbFile);
		if (err == ERROR_SUCCESS && cbFile > 1) {
		    AV *av = newAV();
		    av_push(av, newSVpv(szClsid, cbClsid));
		    av_push(av, newSVpv(szTitle, cbTitle-1));
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

void
_ShowHelpContext(helpfile,context)
    char *helpfile
    IV context
PPCODE:
{
    HWND hwnd;
    dPERINTERP;

    if (!g_hHHCTRL) {
	g_hHHCTRL = LoadLibrary("HHCTRL.OCX");
	if (g_hHHCTRL)
	    g_pfnHtmlHelp = (FNHTMLHELP*)GetProcAddress(g_hHHCTRL, "HtmlHelpA");
    }

    if (!g_pfnHtmlHelp) {
	warn(MY_VERSION ": HtmlHelp control unavailable");
	XSRETURN_EMPTY;
    }

    // HH_HELP_CONTEXT 0x0F: display mapped numeric value in dwData
    hwnd = g_pfnHtmlHelp(GetDesktopWindow(), helpfile, 0x0f, (DWORD)context);

    if (hwnd == 0 && context == 0) // try HH_DISPLAY_TOPIC 0x0
	g_pfnHtmlHelp(GetDesktopWindow(), helpfile, 0, (DWORD)context);
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
	WINOLEOBJECT *pObj = GetOleObject(THIS_ object);
	if (pObj) {
	    HV *olestash = GetWin32OleStash(THIS_ object);
	    SetLastOleError(THIS_ olestash);
	    pEnumObj->pEnum = CreateEnumVARIANT(THIS_ pObj);
	}
    }
    else { /* Clone */
	WINOLEENUMOBJECT *pOriginal = GetOleEnumObject(THIS_ self);
	if (pOriginal) {
	    HV *olestash = GetWin32OleStash(THIS_ self);
	    SetLastOleError(THIS_ olestash);

	    HRESULT hr = pOriginal->pEnum->Clone(&pEnumObj->pEnum);
	    CheckOleError(THIS_ olestash, hr);
	}
    }

    if (!pEnumObj->pEnum) {
	Safefree(pEnumObj);
	XSRETURN_EMPTY;
    }

    AddToObjectChain(THIS_ (OBJECTHEADER*)pEnumObj, WINOLEENUM_MAGIC);

    SV *sv = newSViv((IV)pEnumObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), GetStash(THIS_ self)));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(THIS_ self, TRUE);
    if (pEnumObj) {
	RemoveFromObjectChain(THIS_ (OBJECTHEADER*)pEnumObj);
	if (pEnumObj->pEnum)
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

    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(THIS_ self);
    if (!pEnumObj)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    SV *sv = NULL;
    while (ix == 0 || count-- > 0) {
	sv = NextEnumElement(THIS_ pEnumObj->pEnum, olestash);
	if (!SvOK(sv))
	    break;
	if (!SvIMMORTAL(sv))
	    sv_2mortal(sv);
	if (GIMME_V == G_ARRAY)
	    XPUSHs(sv);
    }

    if (GIMME_V == G_SCALAR && sv && SvOK(sv))
	XPUSHs(sv);
}

void
Reset(self)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(THIS_ self);
    if (!pEnumObj)
	XSRETURN_NO;

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    HRESULT hr = pEnumObj->pEnum->Reset();
    CheckOleError(THIS_ olestash, hr);
    ST(0) = boolSV(hr == S_OK);
    XSRETURN(1);
}

void
Skip(self,...)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(THIS_ self);
    if (!pEnumObj)
	XSRETURN_NO;

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);
    int count = (items > 1) ? SvIV(ST(1)) : 1;
    HRESULT hr = pEnumObj->pEnum->Skip(count);
    CheckOleError(THIS_ olestash, hr);
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
    HRESULT hr;
    WINOLEVARIANTOBJECT *pVarObj;
    VARTYPE vt = items < 2 ? VT_EMPTY : SvIV(ST(1));
    SV *data = items < 3 ? Nullsv : ST(2);

    // XXX Initialize should be superfluous here
    // Initialize();
    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    VARTYPE vt_base = vt & VT_TYPEMASK;
    if (!data && vt_base != VT_NULL && vt_base != VT_EMPTY &&
	vt_base != VT_DISPATCH)
    {
	warn(MY_VERSION ": Win32::OLE::Variant->new(vt, data): data may be"
	     " omitted only for VT_NULL, VT_EMPTY or VT_DISPATCH");
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
	SV *sv = ST(items-1);

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
		if (elt)
		    rgsabound[iDim].lLbound = SvIV(*elt);
		rgsabound[iDim].cElements = 1;
		elt = av_fetch(av, 1, FALSE);
		if (elt)
		    rgsabound[iDim].cElements +=
			SvIV(*elt) - rgsabound[iDim].lLbound;
	    }
	    else
		rgsabound[iDim].cElements = SvIV(sv);
	}

	SAFEARRAY *psa = SafeArrayCreate(vt_base, cDims, rgsabound);
	Safefree(rgsabound);
	if (!psa) {
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
	if (V_ARRAY(pVariant)) {
	    V_VT(pVariant) = VT_UI1 | VT_ARRAY;
	    hr = SafeArrayAccessData(V_ARRAY(pVariant), (void**)&pDest);
	    if (FAILED(hr)) {
		VariantClear(pVariant);
		ReportOleError(THIS_ olestash, hr);
	    }
	    else {
		memcpy(pDest, ptr, len);
		SafeArrayUnaccessData(V_ARRAY(pVariant));
	    }
	}
    }
    else {
	UINT cp = QueryPkgVar(THIS_ olestash, CP_NAME, CP_LEN, cpDefault);
	LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN,
				lcidDefault);
	hr = AssignVariantFromSV(THIS_ data, pVariant, cp, lcid);
	if (FAILED(hr)) {
	    Safefree(pVarObj);
	    ReportOleError(THIS_ olestash, hr);
	    XSRETURN_EMPTY;
	}
    }

    AddToObjectChain(THIS_ (OBJECTHEADER*)pVarObj, WINOLEVARIANT_MAGIC);

    HV *stash = GetStash(THIS_ self);
    SV *sv = newSViv((IV)pVarObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (pVarObj) {
	RemoveFromObjectChain(THIS_ (OBJECTHEADER*)pVarObj);
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
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    HRESULT hr;
    VARIANT variant;
    HV *olestash = GetWin32OleStash(THIS_ self);
    LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN, lcidDefault);

    ST(0) = &PL_sv_undef;
    SetLastOleError(THIS_ olestash);
    VariantInit(&variant);
    hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, type);
    if (SUCCEEDED(hr)) {
	ST(0) = sv_newmortal();
	hr = SetSVFromVariantEx(THIS_ &variant, ST(0), olestash);
    }
    else if (V_VT(&pVarObj->variant) == VT_ERROR) {
	/* special handling for VT_ERROR */
	ST(0) = sv_newmortal();
	V_VT(&variant) = VT_I4;
	V_I4(&variant) = V_ERROR(&pVarObj->variant);
	hr = SetSVFromVariantEx(THIS_ &variant, ST(0), olestash, FALSE);
    }
    VariantClear(&variant);
    CheckOleError(THIS_ olestash, hr);
    XSRETURN(1);
}

void
ChangeType(self,type)
    SV *self
    IV type
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    HRESULT hr = E_INVALIDARG;
    HV *olestash = GetWin32OleStash(THIS_ self);
    LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN, lcidDefault);

    SetLastOleError(THIS_ olestash);
    /* XXX: Does it work with VT_BYREF? */
    hr = VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant,
			     lcid, 0, type);
    CheckOleError(THIS_ olestash, hr);
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
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    HRESULT hr;
    HV *olestash = GetWin32OleStash(THIS_ self);

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

	VARTYPE vt_base = V_VT(pSource) & VT_TYPEMASK;
	V_VT(&variant) = vt_base | VT_BYREF;
	V_VT(&byref) = vt_base;
	if (vt_base == VT_VARIANT)
            V_VARIANTREF(&variant) = &byref;
	else
            V_BYREF(&variant) = &V_BYREF(&byref);

	hr = SafeArrayGetElement(psa, rgIndices, V_BYREF(&variant));
	Safefree(rgIndices);
	if (CheckOleError(THIS_ olestash, hr))
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
	ReportOleError(THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    AddToObjectChain(THIS_ (OBJECTHEADER*)pNewVar, WINOLEVARIANT_MAGIC);

    HV *stash = GetStash(THIS_ self);
    SV *sv = newSViv((IV)pNewVar);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
Date(self,...)
    SV *self
ALIAS:
    Time = 1
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    if (items > 3) {
	char *method[] = {"Date", "Time"};
	warn("Usage: Win32::OLE::Variant::%s"
	      "(SELF [, FORMAT [, LCID]])", method[ix]);
	XSRETURN_EMPTY;
    }

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    char *fmt = NULL;
    DWORD dwFlags = 0;
    LCID lcid = lcidDefault;

    if (items > 1) {
	if (SvIOK(ST(1)))
	    dwFlags = SvIV(ST(1));
	else if SvPOK(ST(1))
	    fmt = SvPV_nolen(ST(1));
    }
    if (items > 2)
	lcid = SvIV(ST(2));
    else
	lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN, lcidDefault);

    HRESULT hr;
    VARIANT variant;
    VariantInit(&variant);
    hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, VT_DATE);
    if (CheckOleError(THIS_ olestash, hr))
        XSRETURN_EMPTY;

    SYSTEMTIME systime;
    VariantTimeToSystemTime(V_DATE(&variant), &systime);

    int len;
    if (ix == 0)
	len = GetDateFormatA(lcid, dwFlags, &systime, fmt, NULL, 0);
    else
	len = GetTimeFormatA(lcid, dwFlags, &systime, fmt, NULL, 0);

    if (len > 1) {
	SV *sv = ST(0) = sv_2mortal(newSV(len));
	if (ix == 0)
	    len = GetDateFormatA(lcid, dwFlags, &systime, fmt, SvPVX(sv), len);
	else
	    len = GetTimeFormatA(lcid, dwFlags, &systime, fmt, SvPVX(sv), len);

	if (len > 1) {
	    SvCUR_set(sv, len-1);
	    SvPOK_on(sv);
	}
    }
    else
        ST(0) = &PL_sv_undef;

    VariantClear(&variant);
    XSRETURN(1);
}

void
Currency(self,...)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    if (items > 3) {
	warn("Usage: Win32::OLE::Variant::Currency"
	      "(SELF [, CURRENCYFMT [, LCID]])");
	XSRETURN_EMPTY;
    }

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    HV *hv = NULL;
    DWORD dwFlags = 0;
    LCID lcid = lcidDefault;

    if (items > 1) {
	SV *format = ST(1);
	if (SvIOK(format))
	    dwFlags = SvIV(format);
	else if (SvROK(format) && SvTYPE(SvRV(format)) == SVt_PVHV)
	    hv = (HV*)SvRV(format);
	else {
	    croak("Win32::OLE::Variant::GetCurrencyFormat: "
		  "CURRENCYFMT must be a HASH reference");
	    XSRETURN_EMPTY;
	}
    }

    if (items > 2)
	lcid = SvIV(ST(2));
    else
	lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN, lcidDefault);

    HRESULT hr;
    VARIANT variant;
    VariantInit(&variant);
    hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, VT_CY);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    CURRENCYFMT fmt;
    Zero(&fmt, 1, CURRENCYFMT);

    fmt.NumDigits        = GetLocaleNumber(THIS_ hv, "NumDigits",
					   lcid, LOCALE_IDIGITS);
    fmt.LeadingZero      = GetLocaleNumber(THIS_ hv, "LeadingZero",
					   lcid, LOCALE_ILZERO);
    fmt.Grouping         = GetLocaleNumber(THIS_ hv, "Grouping",
					   lcid, LOCALE_SMONGROUPING);
    fmt.NegativeOrder    = GetLocaleNumber(THIS_ hv, "NegativeOrder",
					   lcid, LOCALE_INEGCURR);
    fmt.PositiveOrder    = GetLocaleNumber(THIS_ hv, "PositiveOrder",
					   lcid, LOCALE_ICURRENCY);

    fmt.lpDecimalSep     = GetLocaleString(THIS_ hv, "DecimalSep",
					   lcid, LOCALE_SMONDECIMALSEP);
    fmt.lpThousandSep    = GetLocaleString(THIS_ hv, "ThousandSep",
					   lcid, LOCALE_SMONTHOUSANDSEP);
    fmt.lpCurrencySymbol = GetLocaleString(THIS_ hv, "CurrencySymbol",
					   lcid, LOCALE_SCURRENCY);

    int len = 0;
    int sign = 0;
    char amount[40];
    unsigned __int64 u64 = *(unsigned __int64*)&V_CY(&variant);

    if ((__int64)u64 < 0) {
	amount[len++] = '-';
	u64 = (unsigned __int64)(-(__int64)u64);
	sign = 1;
    }
    while (u64) {
	amount[len++] = u64%10 + '0';
	u64 /= 10;
    }
    if (len == sign)
	amount[len++] = '0';
    amount[len] = '\0';
    _strrev(amount+sign);

    /* VT_CY has an implied decimal point before the last 4 digits */
    SV *number;
    if (len-sign < 5)
	number = newSVpvf("%.*s0.%.*s%s", sign, amount,
			  4-(len-sign), "000", amount+sign);
    else
	number = newSVpvf("%.*s.%s", len-4, amount, amount+len-4);

    DBG(("amount='%s' number='%s' len=%d sign=%d", amount, SvPVX(number),
	 len, sign));

    len = GetCurrencyFormatA(lcid, dwFlags, SvPVX(number), &fmt, NULL, 0);
    if (len > 1) {
	SV *sv = ST(0) = sv_2mortal(newSV(len));
	len = GetCurrencyFormatA(lcid, dwFlags, SvPVX(number), &fmt,
				 SvPVX(sv), len);
	if (len > 1) {
	    SvCUR_set(sv, len-1);
	    SvPOK_on(sv);
	}
    }
    else
	ST(0) = &PL_sv_undef;

    SvREFCNT_dec(number);
    VariantClear(&variant);
    XSRETURN(1);
}

void
Number(self,...)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    if (items > 3) {
	warn("Usage: Win32::OLE::Variant::Number"
	      "(SELF [, NUMBERFMT [, LCID]])");
	XSRETURN_EMPTY;
    }

    HV *olestash = GetWin32OleStash(THIS_ self);
    SetLastOleError(THIS_ olestash);

    HV *hv = NULL;
    DWORD dwFlags = 0;
    LCID lcid = lcidDefault;

    if (items > 1) {
	SV *format = ST(1);
	if (SvIOK(format))
	    dwFlags = SvIV(format);
	else if (SvROK(format) && SvTYPE(SvRV(format)) == SVt_PVHV)
	    hv = (HV*)SvRV(format);
	else {
	    croak("Win32::OLE::Variant::GetNumberFormat: "
		  "NUMBERFMT must be a HASH reference");
	    XSRETURN_EMPTY;
	}
    }

    if (items > 2)
	lcid = SvIV(ST(2));
    else
	lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN, lcidDefault);

    HRESULT hr;
    VARIANT variant;
    VariantInit(&variant);
    hr = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, VT_R8);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    NUMBERFMT fmt;
    Zero(&fmt, 1, NUMBERFMT);

    fmt.NumDigits     = GetLocaleNumber(THIS_ hv, "NumDigits",
					lcid, LOCALE_IDIGITS);
    fmt.LeadingZero   = GetLocaleNumber(THIS_ hv, "LeadingZero",
					lcid, LOCALE_ILZERO);
    fmt.Grouping      = GetLocaleNumber(THIS_ hv, "Grouping",
					lcid, LOCALE_SGROUPING);
    fmt.NegativeOrder = GetLocaleNumber(THIS_ hv, "NegativeOrder",
					lcid, LOCALE_INEGNUMBER);

    fmt.lpDecimalSep  = GetLocaleString(THIS_ hv, "DecimalSep",
					lcid, LOCALE_SDECIMAL);
    fmt.lpThousandSep = GetLocaleString(THIS_ hv, "ThousandSep",
					lcid, LOCALE_STHOUSAND);

    SV *number = newSVpvf("%.*f", fmt.NumDigits, V_R8(&variant));
    int len = GetNumberFormatA(lcid, dwFlags, SvPVX(number), &fmt, NULL, 0);
    if (len > 1) {
	SV *sv = ST(0) = sv_2mortal(newSV(len));
	len = GetNumberFormatA(lcid, dwFlags, SvPVX(number), &fmt,
			       SvPVX(sv), len);
	if (len > 1) {
	    SvCUR_set(sv, len-1);
	    SvPOK_on(sv);
	}
    }
    else
	ST(0) = &PL_sv_undef;

    SvREFCNT_dec(number);
    VariantClear(&variant);
    XSRETURN(1);
}

void
Dim(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    VARIANT *pVariant = &pVarObj->variant;
    while (V_VT(pVariant) == (VT_VARIANT | VT_BYREF))
        pVariant = V_VARIANTREF(pVariant);

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

    HV *olestash = GetWin32OleStash(THIS_ self);
    if (CheckOleError(THIS_ olestash, hr))
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
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);
    if (!pVarObj)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(THIS_ self);
    VARIANT *pVariant = &pVarObj->variant;

    while (V_VT(pVariant) == (VT_VARIANT | VT_BYREF))
        pVariant = V_VARIANTREF(pVariant);

    if (!V_ISARRAY(pVariant)) {
	if (items-1 != ix) {
	    warn(MY_VERSION ": Win32::OLE::Variant->%s(): Wrong number of "
		 "arguments" , paszMethod[ix]);
	    XSRETURN_EMPTY;
	}
    scalar_mode:
	HRESULT hr;
	if (ix == 0) { /* Get */
	    ST(0) = sv_newmortal();
	    hr = SetSVFromVariantEx(THIS_ pVariant, ST(0), olestash);
	}
	else { /* Put */
	    UINT cp = QueryPkgVar(THIS_ olestash, CP_NAME, CP_LEN, cpDefault);
	    LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN,
				    lcidDefault);
	    ST(0) = sv_mortalcopy(self);
	    hr = AssignVariantFromSV(THIS_ ST(1), pVariant, cp, lcid);
	}
	CheckOleError(THIS_ olestash, hr);
	XSRETURN(1);
    }

    SAFEARRAY *psa = V_ISBYREF(pVariant) ? *V_ARRAYREF(pVariant)
	                                  : V_ARRAY(pVariant);
    UINT cDims = SafeArrayGetDim(psa);

    /* Special case for one-dimensional VT_UI1 arrays */
    VARTYPE vt_base = V_VT(pVariant) & VT_TYPEMASK;
    if (vt_base == VT_UI1 && cDims == 1 && items-1 == ix)
        goto scalar_mode;

    /* Array Put, e.g. $array->Put([ [11,12], [21,22] ]) */
    if (ix == 1 && items == 2 && SvROK(ST(1)) &&
	SvTYPE(SvRV(ST(1))) == SVt_PVAV)
    {
	UINT cp = QueryPkgVar(THIS_ olestash, CP_NAME, CP_LEN, cpDefault);
	LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN,
				lcidDefault);
	HRESULT hr = SetSafeArrayFromAV(THIS_ (AV*)SvRV(ST(1)), vt_base, psa,
					cDims, cp, lcid);
	CheckOleError(THIS_ olestash, hr);
	ST(0) = sv_mortalcopy(self);
	XSRETURN(1);
    }

    if (items-1 != cDims+ix) {
	warn(MY_VERSION ": Win32::OLE::Variant->%s(): Wrong number of indices; "
	     " dimension of SafeArray is %d", paszMethod[ix], cDims);
	XSRETURN_EMPTY;
    }

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
    else {
        V_BYREF(&variant) = &V_BYREF(&byref);
	if (vt_base == VT_BSTR)
	    V_BSTR(&byref) = NULL;
	else if (vt_base == VT_DISPATCH)
	    V_DISPATCH(&byref) = NULL;
	else if (vt_base == VT_UNKNOWN)
	    V_UNKNOWN(&byref) = NULL;
    }

    HRESULT hr = S_OK;
    if (ix == 0) { /* Get */
	ST(0) = &PL_sv_undef;
	hr = SafeArrayGetElement(psa, rgIndices, V_BYREF(&variant));
	if (SUCCEEDED(hr)) {
	    ST(0) = sv_newmortal();
	    hr = SetSVFromVariantEx(THIS_ &variant, ST(0), olestash);
	}
    }
    else { /* Put */
	UINT cp = QueryPkgVar(THIS_ olestash, CP_NAME, CP_LEN, cpDefault);
	LCID lcid = QueryPkgVar(THIS_ olestash, LCID_NAME, LCID_LEN,
				lcidDefault);
	hr = AssignVariantFromSV(THIS_ ST(items-1), &variant, cp, lcid);
	if (SUCCEEDED(hr)) {
	    if (vt_base == VT_BSTR)
		hr = SafeArrayPutElement(psa, rgIndices, V_BSTR(&byref));
	    else if (vt_base == VT_DISPATCH)
		hr = SafeArrayPutElement(psa, rgIndices, V_DISPATCH(&byref));
	    else if (vt_base == VT_UNKNOWN)
		hr = SafeArrayPutElement(psa, rgIndices, V_UNKNOWN(&byref));
	    else
		hr = SafeArrayPutElement(psa, rgIndices, V_BYREF(&variant));
	}
	if (SUCCEEDED(hr))
	    ST(0) = sv_mortalcopy(self);
    }
    VariantClear(&byref);
    Safefree(rgIndices);
    CheckOleError(THIS_ olestash, hr);
    XSRETURN(1);
}

void
LastError(self,...)
    SV *self
PPCODE:
{
    // Win32::OLE::Variant->LastError() exists only for backward compatibility.
    // It is now just a proxy for Win32::OLE->LastError().

    HV *olestash = GetWin32OleStash(THIS_ self);
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
    _RefType = 3
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj) {
	HRESULT hr;
	HV *olestash = GetWin32OleStash(THIS_ self);
	SetLastOleError(THIS_ olestash);
	ST(0) = sv_newmortal();
	if (ix == 0) /* Type */
	    sv_setiv(ST(0), V_VT(&pVarObj->variant));
	else if (ix == 1) /* Value */
	    hr = SetSVFromVariantEx(THIS_ &pVarObj->variant, ST(0), olestash);
	else if (ix == 2) /* _Value, see also: _Clone (alias of Copy) */
	    hr = SetSVFromVariantEx(THIS_ &pVarObj->variant, ST(0), olestash,
				    TRUE);
	else if (ix == 3)  { /* _RefType */
	    VARIANT *pVariant = &pVarObj->variant;
	    while (V_VT(pVariant) == (VT_BYREF|VT_VARIANT))
		pVariant = V_VARIANTREF(pVariant);
	    sv_setiv(ST(0), V_VT(pVariant));
	}
	CheckOleError(THIS_ olestash, hr);
    }
    XSRETURN(1);
}

void
Unicode(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(THIS_ self);

    ST(0) = &PL_sv_undef;
    if (pVarObj) {
	VARIANT Variant;
	VARIANT *pVariant = &pVarObj->variant;
	HRESULT hr = S_OK;

	HV *olestash = GetWin32OleStash(THIS_ self);
	SetLastOleError(THIS_ olestash);
	VariantInit(&Variant);
	if ((V_VT(pVariant) & ~VT_BYREF) != VT_BSTR) {
	    LCID lcid = QueryPkgVar(THIS_ olestash,
				    LCID_NAME, LCID_LEN, lcidDefault);

	    hr = VariantChangeTypeEx(&Variant, pVariant, lcid, 0, VT_BSTR);
	    pVariant = &Variant;
	}

	if (!CheckOleError(THIS_ olestash, hr)) {
	    BSTR bstr = V_ISBYREF(pVariant) ? *V_BSTRREF(pVariant)
		                            : V_BSTR(pVariant);
	    STRLEN olecharlen = SysStringLen(bstr);
	    SV *sv = newSVpv((char*)bstr, 2*olecharlen);
	    U16 *pus = (U16*)SvPVX(sv);
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
	SvCUR_set(sv, LCMapStringA(lcid, flags, string, length,
				   SvPVX(sv), SvLEN(sv)));
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
	len = GetLocaleInfoA(lcid, lctype, SvPVX(sv), SvLEN(sv));
	if (len) {
	    SvCUR_set(sv, len-1);
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

void
SendSettingChange()
PPCODE:
{
    DWORD dwResult;

    SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, NULL,
		       SMTO_NORMAL, 5000, &dwResult);
    XSRETURN_EMPTY;
}

void
SetLocaleInfo(lcid,lctype,lcdata)
    IV lcid
    IV lctype
    char *lcdata
PPCODE:
{
    if (SetLocaleInfoA(lcid, lctype, lcdata))
	XSRETURN_YES;

    XSRETURN_EMPTY;
}


##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::TypeLib

void
new(self,object)
    SV *self
    SV *object
PPCODE:
{
    HRESULT hr;
    HV *stash = Nullhv;
    ITypeLib *pTypeLib;
    TLIBATTR *pTLibAttr;

    if (sv_isobject(object) && sv_derived_from(object, szWINOLE)) {
	WINOLEOBJECT *pOleObj = GetOleObject(THIS_ object);
	if (!pOleObj)
	    XSRETURN_EMPTY;

	unsigned int count;
	hr = pOleObj->pDispatch->GetTypeInfoCount(&count);
	stash = SvSTASH(pOleObj->self);
	if (CheckOleError(THIS_ stash, hr) || count == 0)
	    XSRETURN_EMPTY;

	ITypeInfo *pTypeInfo;
	hr = pOleObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
	if (CheckOleError(THIS_ stash, hr))
	    XSRETURN_EMPTY;

	unsigned int index;
	hr = pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
	pTypeInfo->Release();
	if (CheckOleError(THIS_ stash, hr))
	    XSRETURN_EMPTY;
    }
    else {
	stash = GetWin32OleStash(THIS_ self);
	UINT cp = QueryPkgVar(THIS_ stash, CP_NAME, CP_LEN, cpDefault);

	char *pszBuffer = SvPV_nolen(object);
	OLECHAR Buffer[OLE_BUF_SIZ];
	OLECHAR *pBuffer = GetWideChar(THIS_ pszBuffer, Buffer, OLE_BUF_SIZ, cp);
	hr = LoadTypeLibEx(pBuffer, REGKIND_NONE, &pTypeLib);
	ReleaseBuffer(THIS_ pBuffer, Buffer);
	if (CheckOleError(THIS_ stash, hr))
	    XSRETURN_EMPTY;
    }

    hr = pTypeLib->GetLibAttr(&pTLibAttr);
    if (FAILED(hr)) {
	pTypeLib->Release();
	ReportOleError(THIS_ stash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeLibObject(THIS_ pTypeLib, pTLibAttr));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (pObj) {
	RemoveFromObjectChain(THIS_ (OBJECTHEADER*)pObj);
	if (pObj->pTypeLib) {
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
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    DWORD dwHelpContext;
    BSTR bstrName, bstrDocString, bstrHelpFile;
    HRESULT hr = pObj->pTypeLib->GetDocumentation(index, &bstrName,
			  &bstrDocString, &dwHelpContext, &bstrHelpFile);
    HV *olestash = GetWin32OleStash(THIS_ self);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = GetDocumentation(THIS_ bstrName, bstrDocString,
			      dwHelpContext, bstrHelpFile);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetLibAttr(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    TLIBATTR *p = pObj->pTLibAttr;
    HV *hv = newHV();

    hv_store(hv, "lcid",          4, newSViv(p->lcid), 0);
    hv_store(hv, "syskind",       7, newSViv(p->syskind), 0);
    hv_store(hv, "wLibFlags",     9, newSViv(p->wLibFlags), 0);
    hv_store(hv, "wMajorVerNum", 12, newSViv(p->wMajorVerNum), 0);
    hv_store(hv, "wMinorVerNum", 12, newSViv(p->wMinorVerNum), 0);
    hv_store(hv, "guid",          4, SetSVFromGUID(THIS_ p->guid), 0);

    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

void
_GetTypeInfoCount(self)
    SV *self
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    XSRETURN_IV(pObj->pTypeLib->GetTypeInfoCount());
}

void
_GetTypeInfo(self,index)
    SV *self
    IV index
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeLib->GetTypeInfo(index, &pTypeInfo);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr));
    XSRETURN(1);
}

void
GetTypeInfo(self,name,...)
    SV *self
    SV *name
PPCODE:
{
    WINOLETYPELIBOBJECT *pObj = GetOleTypeLibObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;

    HV *olestash = GetWin32OleStash(THIS_ self);

    if (SvIOK(name)) {
	HRESULT hr = pObj->pTypeLib->GetTypeInfo(SvIV(name), &pTypeInfo);
	if (CheckOleError(THIS_ olestash, hr))
	    XSRETURN_EMPTY;

	hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(THIS_ olestash, hr);
	    XSRETURN_EMPTY;
	}

	ST(0) = sv_2mortal(CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr));
	XSRETURN(1);
    }

    UINT cp = QueryPkgVar(THIS_ olestash, CP_NAME, CP_LEN, cpDefault);
    TYPEKIND tkind = items > 2 ? (TYPEKIND)SvIV(ST(2)) : TKIND_MAX;
    char *pszName = SvPV_nolen(name);
    int count = pObj->pTypeLib->GetTypeInfoCount();
    for (int index = 0; index < count; ++index) {
	HRESULT hr = pObj->pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (CheckOleError(THIS_ olestash, hr))
	    XSRETURN_EMPTY;

	BSTR bstrName;
	hr = pTypeInfo->GetDocumentation(-1, &bstrName, NULL, NULL, NULL);
	char szStr[OLE_BUF_SIZ];
	char *pszStr = GetMultiByte(THIS_ bstrName, szStr, sizeof(szStr), cp);
	int equal = strEQ(pszStr, pszName);
	ReleaseBuffer(THIS_ pszStr, szStr);
	SysFreeString(bstrName);
	if (!equal) {
	    pTypeInfo->Release();
	    continue;
	}

	hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(hr)) {
	    pTypeInfo->Release();
	    ReportOleError(THIS_ olestash, hr);
	    XSRETURN_EMPTY;
	}

	if (tkind == TKIND_MAX || tkind == pTypeAttr->typekind) {
	    ST(0) = sv_2mortal(CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr));
	    XSRETURN(1);
	}

	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeInfo->Release();
    }
    XSRETURN_EMPTY;
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

    WINOLEOBJECT *pOleObj = GetOleObject(THIS_ object);
    if (!pOleObj)
        XSRETURN_EMPTY;

    unsigned int count;
    HRESULT hr = pOleObj->pDispatch->GetTypeInfoCount(&count);
    HV *olestash = SvSTASH(pOleObj->self);
    if (CheckOleError(THIS_ olestash, hr) || count == 0)
        XSRETURN_EMPTY;

    hr = pOleObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
    if (CheckOleError(THIS_ olestash, hr))
        XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeInfoObject(THIS_ pTypeInfo, pTypeAttr));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (pObj) {
	RemoveFromObjectChain(THIS_ (OBJECTHEADER*)pObj);
	if (pObj->pTypeInfo) {
	    pObj->pTypeInfo->ReleaseTypeAttr(pObj->pTypeAttr);
	    pObj->pTypeInfo->Release();
	}
	Safefree(pObj);
    }
    XSRETURN_EMPTY;
}

void
GetContainingTypeLib(self)
    SV *self
PPCODE:
{
    ITypeLib  *pTypeLib;
    TLIBATTR  *pTLibAttr;

    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    unsigned int index;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetContainingTypeLib(&pTypeLib, &index);
    if (CheckOleError(THIS_ olestash, hr))
        XSRETURN_EMPTY;

    hr = pTypeLib->GetLibAttr(&pTLibAttr);
    if (FAILED(hr)) {
	pTypeLib->Release();
	ReportOleError(THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    ST(0) = sv_2mortal(CreateTypeLibObject(THIS_ pTypeLib, pTLibAttr));
    XSRETURN(1);
}

void
_GetDocumentation(self,memid=-1)
    SV *self
    IV memid
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    DWORD dwHelpContext;
    BSTR bstrName, bstrDocString, bstrHelpFile;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetDocumentation(memid, &bstrName,
			   &bstrDocString, &dwHelpContext, &bstrHelpFile);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = GetDocumentation(THIS_ bstrName, bstrDocString,
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
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    FUNCDESC *p;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetFuncDesc(index, &p);
    if (CheckOleError(THIS_ olestash, hr))
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

    HV *elemdesc = TranslateElemDesc(THIS_ &p->elemdescFunc, pObj, olestash);
    hv_store(hv, "elemdescFunc", 12, newRV_noinc((SV*)elemdesc), 0);

    if (p->cParams > 0) {
	AV *av = newAV();

	for (int i = 0; i < p->cParams; ++i) {
	    elemdesc = TranslateElemDesc(THIS_ &p->lprgelemdescParam[i],
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
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    int flags;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetImplTypeFlags(index, &flags);
    if (CheckOleError(THIS_ olestash, hr))
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

    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetRefTypeOfImplType(index, &hRefType);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pObj->pTypeInfo->GetRefTypeInfo(hRefType, &pTypeInfo);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr)) {
	pTypeInfo->Release();
	ReportOleError(THIS_ olestash, hr);
	XSRETURN_EMPTY;
    }

    New(0, pObj, 1, WINOLETYPEINFOOBJECT);
    pObj->pTypeInfo = pTypeInfo;
    pObj->pTypeAttr = pTypeAttr;

    AddToObjectChain(THIS_ (OBJECTHEADER*)pObj, WINOLETYPEINFO_MAGIC);

    SV *sv = newSViv((IV)pObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), GetStash(THIS_ self)));
    XSRETURN(1);
}

void
_GetNames(self,memid,count)
    SV *self
    IV memid
    IV count
PPCODE:
{
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    BSTR *rgbstr;
    New(0, rgbstr, count, BSTR);
    unsigned int cNames;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetNames(memid, rgbstr, count, &cNames);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    AV *av = newAV();
    for (int i = 0 ; i < cNames ; ++i) {
	char szName[32];
	// XXX use correct codepage ???
	char *pszName = GetMultiByte(THIS_ rgbstr[i],
				     szName, sizeof(szName), CP_ACP);
	SysFreeString(rgbstr[i]);
	av_push(av, newSVpv(pszName, 0));
	ReleaseBuffer(THIS_ pszName, szName);
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
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    TYPEATTR *p = pObj->pTypeAttr;
    HV *hv = newHV();

    hv_store(hv, "guid",              4, SetSVFromGUID(THIS_ p->guid), 0);
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
    WINOLETYPEINFOOBJECT *pObj = GetOleTypeInfoObject(THIS_ self);
    if (!pObj)
	XSRETURN_EMPTY;

    VARDESC *p;
    HV *olestash = GetWin32OleStash(THIS_ self);
    HRESULT hr = pObj->pTypeInfo->GetVarDesc(index, &p);
    if (CheckOleError(THIS_ olestash, hr))
	XSRETURN_EMPTY;

    HV *hv = newHV();
    hv_store(hv, "memid",        5, newSViv(p->memid), 0);
    // LPOLESTR lpstrSchema;
    hv_store(hv, "wVarFlags",    9, newSViv(p->wVarFlags), 0);
    hv_store(hv, "varkind",      7, newSViv(p->varkind), 0);

    HV *elemdesc = TranslateElemDesc(THIS_ &p->elemdescVar,
				     pObj, olestash);
    hv_store(hv, "elemdescVar", 11, newRV_noinc((SV*)elemdesc), 0);

    if (p->varkind == VAR_PERINSTANCE)
	hv_store(hv, "oInst",    5, newSViv(p->oInst), 0);

    if (p->varkind == VAR_CONST) {
	// XXX should be stored as a Win32::OLE::Variant object ?
	SV *sv = newSV(0);
	SetSVFromVariantEx(THIS_ p->lpvarValue, sv, olestash);
	hv_store(hv, "varValue", 8, sv, 0);
    }

    pObj->pTypeInfo->ReleaseVarDesc(p);
    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}
