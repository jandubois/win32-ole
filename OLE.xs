/* OLE.xs
 *
 *  (c) 1995 Microsoft Corporation. All rights reserved.
 *  Developed by ActiveWare Internet Corp., http://www.ActiveWare.com
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
 * - Package Win32::OLE           Constructor and method invokation
 * - Package Win32::OLE::Tie      Implements properties as tied hash
 * - Package Win32::OLE::Const    Load application constants from type library
 * - Package Win32::OLE::Enum     OLE collection enumeration
 * - Package Win32::OLE::Variant  Implements Perl VARIANT objects
 *
 */

#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#include <windows.h>
#ifdef _DEBUG
    #include <crtdbg.h>
    #define DEBUGBREAK _CrtDbgBreak()
#else
    #define DEBUGBREAK
#endif

#if defined(__cplusplus)
extern "C" {
#endif

#include "EXTERN.h"
#include "perl.h"
#include "XSub.h"

#if !defined(_DEBUG)
#    define DBG(a)
#else
#    define DBG(a)  MyDebug a
void
MyDebug(const char *pat, ...)
{
    va_list args;
    va_start(args, pat);
    PerlIO_vprintf(PerlIO_stderr(), pat, args);
    va_end(args);
    PerlIO_flush(PerlIO_stderr());
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

/* common object header */
typedef struct _tagOBJECTHEADER OBJECTHEADER;
typedef struct _tagOBJECTHEADER
{
    long lMagic;
    OBJECTHEADER *pNext;
    OBJECTHEADER *pPrevious;

}   OBJECTHEADER;

/* Win32::OLE object */
typedef struct
{
    OBJECTHEADER header;

    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo;
    IEnumVARIANT *pEnum;

    HV *stash;
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

}   WINOLEVARIANTOBJECT;

/* global variables */
static CRITICAL_SECTION CriticalSection;
static OBJECTHEADER *g_pObj = NULL;

/* forward declarations */
HRESULT SetSVFromVariant(VARIANTARG *pVariant, SV* sv, HV *stash);

/* The following strategy is used to avoid the limitations of hardcoded
 * buffer sizes: Conversion between wide char and multibyte strings
 * is performed by GetMultiByte and GetWideChar respectively. The
 * caller passes a default buffer and size. If the buffer is too small
 * then the conversion routine allocates a new buffer that is big enough.
 * The caller must free this buffer using the ReleaseBuffer function. */

inline void
ReleaseBuffer(void *pszHeap, void *pszStack)
{
    if (pszHeap != pszStack)
	Safefree(pszHeap);
}

char *
GetMultiByte(OLECHAR *wide, char *psz, int len, UINT cp)
{
    *psz = (char) 0;
    if (wide == NULL)
	return psz;

    int count = WideCharToMultiByte(cp, 0, wide, -1, psz, len, NULL, NULL);
    if (count > 0)
	return psz;

    count = WideCharToMultiByte(cp, 0, wide, -1, NULL, 0, NULL, NULL);
    if (count == 0) { /* should never happen */
	warn("Win32::OLE: GetMultiByte() failure: %lu", GetLastError());
	DEBUGBREAK;
	return psz;
    }

    New(0, psz, count, char);
    count = WideCharToMultiByte(cp, 0, wide, -1, psz, count, NULL, NULL);
    return psz;
}

OLECHAR *
GetWideChar(char *psz, OLECHAR *wide, int len, UINT cp)
{
    /* Note: len is number of OLECHARs, not bytes! */
    int count = MultiByteToWideChar(cp, 0, psz, -1, wide, len);
    if (count > 0)
	return wide;

    count = MultiByteToWideChar(cp, 0, psz, -1, NULL, 0);
    if (count == 0) {
	warn("Win32::OLE: GetWideChar() failure: %lu", GetLastError());
	DEBUGBREAK;
	*wide = (OLECHAR) 0;
	return wide;
    }

    New(0, wide, count, OLECHAR);
    count = MultiByteToWideChar(cp, 0, psz, -1, wide, count);
    return wide;
}

IV
QueryPkgVar(HV *stash, char *var, STRLEN len, IV def)
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
ReportOleError(HV *stash, HRESULT res, EXCEPINFO *pExcepInfo, SV *svDetails)
{
    dSP;

    /* Find $Win32::OLE::LastError */
    SV *sv = sv_2mortal(newSVpv(HvNAME(stash), 0));
    sv_catpvn(sv, "::", 2);
    sv_catpvn(sv, LASTERR_NAME, LASTERR_LEN);
    SV *lasterr = perl_get_sv(SvPV(sv, na), TRUE);
    if (lasterr == NULL) {
	warn("Win32::OLE: ReportOleError: couldnot create package variable %s",
	     LASTERR_NAME);
	DEBUGBREAK;
    }

    IV warn = QueryPkgVar(stash, WARN_NAME, WARN_LEN, 0);

    SvCUR_set(sv, 0);

    /* start with exception info */
    if (pExcepInfo != NULL && (pExcepInfo->bstrSource != NULL ||
			       pExcepInfo->bstrDescription != NULL )) {
	char szSource[80] = "<Unknown Source>";
	char szDescription[200] = "<No description provided>";

	char *pszSource = szSource;
	char *pszDescription = szDescription;

	UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

	if (pExcepInfo->bstrSource != NULL)
	    pszSource = GetMultiByte(pExcepInfo->bstrSource, szSource,
				     sizeof(szSource), cp);

	if (pExcepInfo->bstrDescription != NULL)
	    pszDescription = GetMultiByte(pExcepInfo->bstrDescription,
			szDescription, sizeof(szDescription), cp);

	sv_setpvf(sv, "OLE exception from \"%s\":\n\n%s\n\n",
		  pszSource, pszDescription);

	ReleaseBuffer(pszSource, szSource);
	ReleaseBuffer(pszDescription, szDescription);
	/* SysFreeString accepts NULL too */
	SysFreeString(pExcepInfo->bstrSource);
	SysFreeString(pExcepInfo->bstrDescription);
	SysFreeString(pExcepInfo->bstrHelpFile);
    }

    /* always include OLE error code */
    sv_catpvf(sv, "OLE error 0x%08x", res);

    /* try to append ': "error text"' from message catalog */
    char *pszMsgText;
    DWORD dwCount = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
				  FORMAT_MESSAGE_FROM_SYSTEM |
				  FORMAT_MESSAGE_IGNORE_INSERTS,
				  NULL, res, lcidSystemDefault,
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
    if (svDetails != NULL) {
	sv_catpv(sv, "\n    ");
	sv_catsv(sv, svDetails);
    }

    /* try to keep linelength of description below 80 chars. */
    char *pLastBlank = NULL;
    char *pch = SvPV(sv, na);
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

    if (lasterr != NULL) {
	sv_setiv(lasterr, (IV)res);
	sv_setpv(lasterr, SvPVX(sv));
	SvIOK_on(lasterr);
    }

    if (warn > 1 || (warn == 1 && dowarn)) {
	PUSHMARK(sp) ;
	XPUSHs(sv);
	PUTBACK;
	perl_call_pv(warn < 3 ? "Carp::carp" : "Carp::croak", G_DISCARD);
    }

}   /* ReportOleError */

inline BOOL
CheckOleError(HV *stash, HRESULT res, EXCEPINFO *pExcepInfo, SV *svDetails)
{
    if (FAILED(res)) {
	ReportOleError(stash, res, pExcepInfo, svDetails);
	return TRUE;
    }
    return FALSE;
}

SV *
CheckDestroyFunction(SV *sv, char *szMethod)
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
AddToObjectChain(OBJECTHEADER *pHeader, long lMagic)
{
    EnterCriticalSection(&CriticalSection);
    pHeader->lMagic = lMagic;
    pHeader->pPrevious = NULL;
    pHeader->pNext = g_pObj;
    if (g_pObj)
	g_pObj->pPrevious = pHeader;
    g_pObj = pHeader;
    LeaveCriticalSection(&CriticalSection);
}

void
RemoveFromObjectChain(OBJECTHEADER *pHeader)
{
    if (pHeader == NULL)
	return;

    EnterCriticalSection(&CriticalSection);
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
    LeaveCriticalSection(&CriticalSection);
}

SV *
CreatePerlObject(HV *stash, IDispatch *pDispatch, SV *destroy)
{
    /* returns a mortal reference to a new Perl OLE object */

    if (pDispatch == NULL) {
	warn("Win32::OLE: CreatePerlObject() No IDispatch interface");
	DEBUGBREAK;
	return &sv_undef;
    }

    WINOLEOBJECT *pObj;
    HV *hvouter = newHV();
    HV *hvinner = newHV();
    SV *inner;
    SV *sv;
    GV **gv = (GV **) hv_fetch(stash, TIE_NAME, TIE_LEN, FALSE);
    char *szTie = szWINOLETIE;

    if (gv != NULL && (sv = GvSV(*gv)) != NULL && SvPOK(sv))
	szTie = SvPV(sv, na);

    New(0, pObj, 1, WINOLEOBJECT);
    pObj->pDispatch = pDispatch;
    pObj->pTypeInfo = NULL;
    pObj->pEnum = NULL;
    pObj->hashTable = newHV();
    pObj->stash = stash;
    SvREFCNT_inc(stash);

    pObj->destroy = NULL;
    if (destroy !=NULL) {
	if (SvPOK(destroy))
	    pObj->destroy = newSVsv(destroy);
	else if (SvROK(destroy) && SvTYPE(SvRV(destroy)) == SVt_PVCV)
	    pObj->destroy = newRV_inc(SvRV(destroy));
    }

    AddToObjectChain(&pObj->header, WINOLE_MAGIC);


    DBG(("CreatePerlObject = |%lx| Class = %s Tie = %s\n", pObj, 
	 HvNAME(stash), szTie));

    hv_store(hvinner, PERL_OLE_ID, PERL_OLE_IDLEN, newSViv((long)pObj), 0);
    inner = sv_bless(newRV_noinc((SV*)hvinner), gv_stashpv(szTie, TRUE));
    sv_magic((SV*)hvouter, inner, 'P', Nullch, 0);
    SvREFCNT_dec(inner);

    return sv_2mortal(sv_bless(newRV_noinc((SV*)hvouter), stash));

}   /* CreatePerlObject */

void
DestroyPerlObject(WINOLEOBJECT *pObj)
{
    /* Note: Code to clean up external resources should be
     * duplicated in the DllMain/DllEntryPoint function below.
     */

    RemoveFromObjectChain((OBJECTHEADER *)pObj);

    if (pObj->pDispatch == NULL) {
	warn("Win32::OLE: DestroyPerlObject() Object already destroyed");
	DEBUGBREAK;
	/* should never happen, so don't try to do further cleanup */
	return;
    }

    DBG(("DestroyPerlObject |%lx|", pObj));

    DBG((" pDispatch"));
    pObj->pDispatch->Release();
    pObj->pDispatch = NULL;

    DBG((" stash(%d)", SvREFCNT(pObj->stash)));
    SvREFCNT_dec(pObj->stash);

    DBG((" hashTable(%d)", SvREFCNT(pObj->hashTable)));
    SvREFCNT_dec(pObj->hashTable);
    pObj->hashTable = NULL;

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

    DBG(("\n"));
    Safefree(pObj);

}   /* DestroyPerlObject */

WINOLEOBJECT *
CheckOleStruct(IV addr)
{
    WINOLEOBJECT *pObj = (WINOLEOBJECT*)addr;

    if (pObj == NULL || pObj->header.lMagic != WINOLE_MAGIC) {
	warn("Win32::OLE: CheckOleStruct() Not a %s object", szWINOLE);
	DEBUGBREAK;
	pObj = NULL;
    }

    return pObj;
}

WINOLEOBJECT *
GetOleObject(SV *sv)
{
    /*  don't use sv_isobject/sv_derived_from; they'll call mg_get! */
    if (sv != NULL && SvROK(sv) && SvTYPE(SvRV(sv)) == SVt_PVHV) {
	SV **psv = hv_fetch((HV*)SvRV(sv), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
	if (psv != NULL) {
	    IV addr = SvIV(*psv);
	    DBG(("GetOleObject = |%lx|\n", addr));
	    return CheckOleStruct(addr);
	}
    }
    warn("Win32::OLE: GetOleObject() Not a %s object", szWINOLE);
    DEBUGBREAK;
    return (WINOLEOBJECT*)NULL;
}

WINOLEENUMOBJECT *
GetOleEnumObject(SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLEENUM)) {
	WINOLEENUMOBJECT *pEnumObj = (WINOLEENUMOBJECT*)SvIV(SvRV(sv));
	if (pEnumObj != NULL && pEnumObj->header.lMagic == WINOLEENUM_MAGIC)
	    return pEnumObj;
    }
    warn("Win32::OLE: GetOleEnumObject() Not a %s object", szWINOLEENUM);
    DEBUGBREAK;
    return (WINOLEENUMOBJECT*)NULL;
}

WINOLEVARIANTOBJECT *
GetOleVariantObject(SV *sv)
{
    if (sv_isobject(sv) && sv_derived_from(sv, szWINOLEVARIANT)) {
	WINOLEVARIANTOBJECT *pVarObj = (WINOLEVARIANTOBJECT*)SvIV(SvRV(sv));
	if (pVarObj != NULL && pVarObj->header.lMagic == WINOLEVARIANT_MAGIC)
	    return pVarObj;
    }
    warn("Win32::OLE: GetOleVariantObject() Not a %s object", szWINOLEVARIANT);
    DEBUGBREAK;
    return (WINOLEVARIANTOBJECT*)NULL;
}

BSTR
AllocOleString(char* pStr, int length, UINT cp)
{
    int count = MultiByteToWideChar(cp, 0, pStr, length, NULL, 0);
    BSTR bstr = SysAllocStringLen(NULL, count);
    MultiByteToWideChar(cp, 0, pStr, length, bstr, count);
    return bstr;
}

HRESULT
GetHashedDispID(WINOLEOBJECT *pObj, char *buffer, STRLEN len,
		DISPID &dispID, LCID lcid, UINT cp)
{
    HRESULT res;

    if (len == 0 || *buffer == '\0') {
	dispID = DISPID_VALUE;
	return TRUE;
    }

    SV **psv = hv_fetch(pObj->hashTable, buffer, len, 0);
    if (psv != NULL) {
	dispID = (DISPID)SvIV(*psv);
	return TRUE;
    }

    /* not there so get info and add it */
    DISPID id;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;

    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ, cp);
    res = pObj->pDispatch->GetIDsOfNames(IID_NULL, &pBuffer, 1, lcid, &id);
    ReleaseBuffer(pBuffer, Buffer);
    /* Don't call CheckOleError! Caller might retry the "unnamed" method */
    if (SUCCEEDED(res)) {
	hv_store(pObj->hashTable, buffer, len, newSViv(id), 0);
	dispID = id;
    }
    return res;

}   /* GetHashedDispID */

void
FetchTypeInfo(WINOLEOBJECT *pObj)
{
    unsigned int count;
    ITypeInfo *pTypeInfo;
    LPTYPEATTR pTypeAttr;

    if (pObj->pTypeInfo != NULL)
	return;

    HRESULT res = pObj->pDispatch->GetTypeInfoCount(&count);
    if (res == E_NOTIMPL || count == 0) {
	DBG(("GetTypeInfoCount returned %u (count=%d)", res, count));
	return;
    }

    if (CheckOleError(pObj->stash, res, NULL, NULL)) {
	warn("Win32::OLE: FetchTypeInfo() GetTypeInfoCount failed");
	DEBUGBREAK;
	return;
    }

    LCID lcid = QueryPkgVar(pObj->stash, LCID_NAME, LCID_LEN, lcidDefault);
    res = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(pObj->stash, res, NULL, NULL))
	return;

    res = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(res)) {
	pTypeInfo->Release();
	ReportOleError(pObj->stash, res, NULL, NULL);
	return;
    }

    if (pTypeAttr->typekind != TKIND_DISPATCH) {
	int cImplTypes = pTypeAttr->cImplTypes;
	pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	pTypeAttr = NULL;

	for (int i=0 ; i < cImplTypes ; ++i) {
	    HREFTYPE hreftype;
	    ITypeInfo *pRefTypeInfo;

	    res = pTypeInfo->GetRefTypeOfImplType(i, &hreftype);
	    if (FAILED(res))
		break;

	    res = pTypeInfo->GetRefTypeInfo(hreftype, &pRefTypeInfo);
	    if (FAILED(res))
		break;

	    res = pRefTypeInfo->GetTypeAttr(&pTypeAttr);
	    if (FAILED(res)) {
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

    if (FAILED(res)) {
	pTypeInfo->Release();
	ReportOleError(pObj->stash, res, NULL, NULL);
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
NextPropertyName(WINOLEOBJECT *pObj)
{
    HRESULT res;
    unsigned int cName;
    BSTR bstr;
    char szName[64];

    if (pObj->pTypeInfo == NULL)
	return &sv_undef;

    UINT cp = QueryPkgVar(pObj->stash, CP_NAME, CP_LEN, cpDefault);

    while (pObj->PropIndex < pObj->cFuncs+pObj->cVars) {
	ULONG index = pObj->PropIndex++;
	/* Try all the INVOKE_PROPERTYGET functions first */
	if (index < pObj->cFuncs) {
	    LPFUNCDESC pFuncDesc;

	    res = pObj->pTypeInfo->GetFuncDesc(index, &pFuncDesc);
	    if (CheckOleError(pObj->stash, res, NULL, NULL))
		continue;

	    if (!(pFuncDesc->funckind & FUNC_DISPATCH) ||
		!(pFuncDesc->invkind & INVOKE_PROPERTYGET) ||
	        (pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED |
					  FUNCFLAG_FHIDDEN |
					  FUNCFLAG_FNONBROWSABLE))) {
		pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		continue;
	    }

	    res = pObj->pTypeInfo->GetNames(pFuncDesc->memid, &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
	    if (CheckOleError(pObj->stash, res, NULL, NULL)
		|| cName == 0 || bstr == NULL)
		continue;

	    char *pszName = GetMultiByte(bstr, szName, sizeof(szName), cp);
	    SV *sv = newSVpv(pszName, 0);
	    SysFreeString(bstr);
	    ReleaseBuffer(pszName, szName);
	    return sv;
	}
	/* Now try the VAR_DISPATCH kind variables used by older OLE versions */
	else {
	    LPVARDESC pVarDesc;

	    index -= pObj->cFuncs;
	    res = pObj->pTypeInfo->GetVarDesc(index, &pVarDesc);
	    if (CheckOleError(pObj->stash, res, NULL, NULL))
		continue;

	    if (!(pVarDesc->varkind & VAR_DISPATCH) ||
		(pVarDesc->wVarFlags & (VARFLAG_FRESTRICTED |
					VARFLAG_FHIDDEN |
					VARFLAG_FNONBROWSABLE))) {
		pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
		continue;
	    }

	    res = pObj->pTypeInfo->GetNames(pVarDesc->memid, &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
	    if (CheckOleError(pObj->stash, res, NULL, NULL)
		|| cName == 0 || bstr == NULL)
		continue;

	    char *pszName = GetMultiByte(bstr, szName, sizeof(szName), cp);
	    SV *sv = newSVpv(pszName, 0);
	    SysFreeString(bstr);
	    ReleaseBuffer(pszName, szName);
	    return sv;
	}
    }
    return &sv_undef;

}   /* NextPropertyName */

IEnumVARIANT *
CreateEnumVARIANT(WINOLEOBJECT *pObj)
{
    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    HRESULT res;
    IUnknown *punk;
    IEnumVARIANT *pEnum = NULL;

    VariantInit(&result);
    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = 0;

    LCID lcid = QueryPkgVar(pObj->stash, LCID_NAME, LCID_LEN, lcidDefault);

    Zero(&excepinfo, 1, EXCEPINFO);
    res = pObj->pDispatch->Invoke(DISPID_NEWENUM, IID_NULL,
			    lcid, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			    &dispParams, &result, &excepinfo, &argErr);
    if (CheckOleError(pObj->stash, res, &excepinfo, NULL) ||
	(V_VT(&result)&~VT_BYREF) != VT_UNKNOWN)
	return NULL;;

    if (V_ISBYREF(&result))
	punk = *V_UNKNOWNREF(&result);
    else
	punk = V_UNKNOWN(&result);

    res = punk->QueryInterface(IID_IEnumVARIANT, (void**)&pEnum);
    VariantClear(&result);
    CheckOleError(pObj->stash, res, NULL, NULL);
    return pEnum;

}   /* CreateEnumVARIANT */

SV *
NextEnumElement(IEnumVARIANT *pEnum, HV *stash)
{
    SV *sv = &sv_undef;
    VARIANT variant;

    VariantInit(&variant);
    if (pEnum->Next(1, &variant, NULL) == S_OK) {
	sv = newSVpv("",0);
	HRESULT res = SetSVFromVariant(&variant, sv, stash);
	VariantClear(&variant);
	CheckOleError(stash, res, NULL, NULL);
    }
    return sv;

}   /* NextEnumElement */

HRESULT
SetVariantFromSV(SV* sv, VARIANT *pVariant, UINT cp)
{
    HRESULT res = S_OK;
    VariantInit(pVariant);

    /* XXX requirement to call mg_get() may change in Perl > 5.004 */
    if (SvGMAGICAL(sv))
	mg_get(sv);

    /* Objects */
    if (SvROK(sv)) {
	if (sv_derived_from(sv, szWINOLE)) {
	    WINOLEOBJECT *pObj = GetOleObject(sv);
	    if (pObj == NULL)
		res = E_POINTER;
	    else {
		pObj->pDispatch->AddRef();
		V_VT(pVariant) = VT_DISPATCH;
		V_DISPATCH(pVariant) = pObj->pDispatch;
	    }
	    return res;
	}

	if (sv_derived_from(sv, szWINOLEVARIANT)) {
	    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(sv);
	    if (pVarObj == NULL)
		res = E_POINTER;
	    else
		res = VariantCopy(pVariant, &pVarObj->variant);

	    return res;
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
	V_ARRAY(pVariant) = SafeArrayCreate(VT_VARIANT, dim, psab);
	if (V_ARRAY(pVariant) == NULL)
	    res = E_OUTOFMEMORY;
	else {
	    V_VT(pVariant) = VT_VARIANT | VT_ARRAY;

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
			VARIANT variant;
			res = SetVariantFromSV(*psv, &variant, cp);
			if (SUCCEEDED(res)) {
			    res = SafeArrayPutElement(V_ARRAY(pVariant),
						      pix, &variant);
			    VariantClear(&variant);
			}
			if (FAILED(res)) {
			    VariantClear(pVariant);
			    break;
			}
		    }
		}

		while (index >= 0) {
		    if (++pix[index] < plen[index])
			break;
		    pix[index--] = 0;
		}
	    }
	}

	Safefree(pav);
	Safefree(pix);
	Safefree(plen);
	Safefree(psab);

	return res;
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
	V_BSTR(pVariant) = AllocOleString(ptr, len, cp);
    }
    else {
	V_VT(pVariant) = VT_ERROR;
	V_ERROR(pVariant) = DISP_E_PARAMNOTFOUND;
    }

    return res;

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
SetSVFromVariant(VARIANTARG *pVariant, SV* sv, HV *stash)
{
    HRESULT res = S_OK;
    sv_setsv(sv, &sv_undef);

    if (V_ISARRAY(pVariant)) {
	SAFEARRAY *psa = V_ARRAY(pVariant);
	AV **pav;
	VARIANT variant;
	void *pData = &variant;
	IV index;
	long *pArrayIndex, *pLowerBound, *pUpperBound;

	int dim = SafeArrayGetDim(psa);

	VariantInit(&variant);
	V_VT(&variant) = V_VT(pVariant) & ~VT_ARRAY;
	if (V_VT(&variant) != VT_VARIANT)
	    pData = &V_UI1(&variant);

	/* convert 1-dim UI1 ARRAY to simple SvPV */
	if (dim == 1 && V_VT(&variant) == VT_UI1) {
	    char *pStr;
	    long lLower, lUpper;

	    SafeArrayGetLBound(psa, 1, &lLower);
	    SafeArrayGetUBound(psa, 1, &lUpper);
	    res = SafeArrayAccessData(psa, (void**)&pStr);
	    if (SUCCEEDED(res)) {
		sv_setpvn(sv, pStr, lUpper-lLower+1);
		SafeArrayUnaccessData(psa);
	    }

	    return res;
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

	while (index >= 0) {
	    res = SafeArrayGetElement(psa, pArrayIndex, pData);
	    if (FAILED(res))
		break;

	    SV *val = newSVpv("",0);
	    res = SetSVFromVariant(&variant, val, stash);
	    VariantClear(&variant);
	    if (FAILED(res)) {
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

	for (index = 1 ; index < dim ; ++index)
	    SvREFCNT_dec((SV*)pav[index]);

	if (FAILED(res))
	    SvREFCNT_dec((SV*)*pav);
	else {
	    SV *retval = newRV_noinc((SV*)*pav);

	    /* eliminate all outer-level single-element lists */
	    //while (SvROK(retval)) {
	    //	AV *av = (AV*)SvRV(retval);
	    //	if (SvTYPE((SV*)av) != SVt_PVAV || av_len(av) != 0)
	    //	    break;
	    //	SV *temp = av_pop(av);
	    //	SvREFCNT_dec(retval);
	    //	retval = temp;
	    //}

	    sv_setsv(sv, retval);
	    SvREFCNT_dec(retval);
	}

	Safefree(pArrayIndex);
	Safefree(pLowerBound);
	Safefree(pUpperBound);
	Safefree(pav);

	return res;
    }

    switch(V_VT(pVariant) & ~VT_BYREF)
    {
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
ConvertString:
    {
	char Str[260];
	char *pStr;
	UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

	if (V_ISBYREF(pVariant))
	    pStr = GetMultiByte(*V_BSTRREF(pVariant), Str, sizeof(Str), cp);
	else
	    pStr = GetMultiByte(V_BSTR(pVariant), Str, sizeof(Str), cp);

	sv_setpv(sv, pStr);
	ReleaseBuffer(pStr, Str);
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
	    sv_setsv(sv, CreatePerlObject(stash, pDispatch, NULL));
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
	    SUCCEEDED(punk->QueryInterface(IID_IDispatch,
					   (void**)&pDispatch)))
	{
	    sv_setsv(sv, CreatePerlObject(stash, pDispatch, NULL));
	}
	break;
    }

    case VT_DATE:
    case VT_CY:
    case VT_VARIANT:
    default:
    {
	LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);
	HRESULT res = VariantChangeTypeEx(pVariant, pVariant, lcid, 0, VT_BSTR);
	if (SUCCEEDED(res))
	    goto ConvertString;
	break;
    }
    }

    return res;

}   /* SetSVFromVariant */

BOOL APIENTRY
#ifdef __BORLANDC__
DllEntryPoint
#else
DllMain
#endif
(HANDLE hModule, DWORD fdwReason, LPVOID lpvReserved)
{
    switch (fdwReason) {
    case DLL_PROCESS_ATTACH:
	InitializeCriticalSection(&CriticalSection);
	OleInitialize(NULL);
	break;

    case DLL_PROCESS_DETACH:
	/* Global destruction will have normally DESTROYed all
	 * objects, so the loop below will never be entered.
	 * Unless global destruction phase was somehow interrupted.
	 * Only external resources are cleaned up here.
	 */

	/* XXX Should we EnterCriticalSection(&CriticalSection) ??? */
	DBG(("DLL_PROCESS_DETACH\n"));

	while (g_pObj != NULL) {
	    DBG(("Cleaning out escaped object |%lx|\n", g_pObj));

	    switch (g_pObj->lMagic)
	    {
	    case WINOLE_MAGIC:
	    {
		WINOLEOBJECT *pObj = (WINOLEOBJECT*)g_pObj;
		if (pObj->pDispatch != NULL)
		    pObj->pDispatch->Release();
		if (pObj->pTypeInfo != NULL)
		    pObj->pTypeInfo->Release();
		if (pObj->pEnum != NULL)
		    pObj->pEnum->Release();
		break;
	    }

	    case WINOLEENUM_MAGIC:
	    {
		WINOLEENUMOBJECT *pEnumObj = (WINOLEENUMOBJECT*)g_pObj;
		if (pEnumObj->pEnum != NULL)
		    pEnumObj->pEnum->Release();
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

	DBG(("OleUninitialize\n"));
	OleUninitialize();
	DBG(("DeleteCriticalSection\n"));
	DeleteCriticalSection(&CriticalSection);
	DBG(("Really the end...\n"));
	break;
    }

    return TRUE;

}   /* DllMain/DllEntryPoint */

BOOL
CallObjectMethod(SV **mark, I32 ax, I32 items, char *pszMethod)
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
GetStash(SV *sv)
{
    if (sv_isobject(sv))
	return SvSTASH(SvRV(sv));
    else if (SvPOK(sv))
	return gv_stashsv(sv, TRUE);
    else
	return (HV *)&sv_undef;
}

#if defined(__cplusplus)
}
#endif

/*##########################################################################*/

MODULE = Win32::OLE		PACKAGE = Win32::OLE

PROTOTYPES: DISABLE

void
new(...)
PPCODE:
{
    UINT cp;
    CLSID CLSIDObj;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    unsigned int length;
    char *buffer;
    HKEY handle;
    IDispatch *pDispatch;
    HRESULT res;

    if (CallObjectMethod(mark, ax, items, "new"))
	return;

    if (items < 2 || items > 3) {
	warn("Usage: Win32::OLE->new(class[,destroy])");
	DEBUGBREAK;
	XSRETURN_EMPTY;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *oleclass = ST(1);
    SV *destroy = NULL;

    ST(0) = &sv_undef;

    if (items == 3)
	destroy = CheckDestroyFunction(ST(2), "Win32::OLE::new");

    cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);
    buffer = SvPV(oleclass, length);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ, cp);
    res = CLSIDFromProgID(pBuffer, &CLSIDObj);
    ReleaseBuffer(pBuffer, Buffer);

    if (!CheckOleError(stash, res, NULL, NULL)) {
	res = CoCreateInstance(CLSIDObj, NULL, CLSCTX_LOCAL_SERVER,
			       IID_IDispatch, (void**)&pDispatch);
	if (FAILED(res)) {
	    res = CoCreateInstance(CLSIDObj, NULL, CLSCTX_ALL,
				   IID_IDispatch, (void**)&pDispatch);
	    CheckOleError(stash, res, NULL, NULL);
	}

	if (SUCCEEDED(res)) {
	    ST(0) = CreatePerlObject(stash, pDispatch, destroy);
	    DBG(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
	}
    }
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(self);

    DBG(("Win32::OLE::DESTROY |%lx|\n", pObj));
    if (pObj != NULL && pObj->destroy != NULL) {
	if (SvPOK(pObj->destroy)) {
	    /* Dispatch($self,$destroy,$retval); */
	    EXTEND(sp,2);
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
    VARIANT result;
    DISPPARAMS dispParams;
    SV *curitem, *sv;
    HE **rghe = NULL; /* named argument names */

    SV *err = NULL; /* error details */
    HRESULT res = S_OK;

    VariantInit(&result);
    ST(0) = &sv_no;

    if (!sv_isobject(self)) {
	warn("Win32::OLE::Dispatch: Cannot be called as class method");
	DEBUGBREAK;
	XSRETURN(1);
    }

    pObj = GetOleObject(self);
    if (pObj == NULL) {
	XSRETURN(1);
    }

    LCID lcid = QueryPkgVar(pObj->stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(pObj->stash, CP_NAME, CP_LEN, cpDefault);

    if (SvPOK(method)) {
	buffer = SvPV(method, length);
	if (length > 0) {
	    res = GetHashedDispID(pObj, buffer, length, dispID, lcid, cp);
	    if (FAILED(res)) {
		err = sv_2mortal(newSVpvf(" in GetIDsOfNames \"%s\"", buffer));
		ReportOleError(pObj->stash, res, NULL, err);
		ST(0) = &sv_undef;
		XSRETURN(1);
	    }
	}
    }

    DBG(("Dispatch \"%s\"\n", buffer));

    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = items - 3;

    Zero(&excepinfo, 1, EXCEPINFO);
    VariantInit(&result);

    /* last arg is ref to a non-object-hash => named arguments */
    curitem = ST(items-1);
    if (SvROK(curitem) && (sv = SvRV(curitem)) &&
	SvTYPE(sv) == SVt_PVHV && !SvOBJECT(sv))
    {
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

	rgszNames[0] = AllocOleString(buffer, length, cp);
	hv_iterinit(hv);
	for (index = 0; index < dispParams.cNamedArgs; ++index) {
	    rghe[index] = hv_iternext(hv);
	    char *pszName = hv_iterkey(rghe[index], &len);
	    rgszNames[1+index] = AllocOleString(pszName, len, cp);
	}

	res = pObj->pDispatch->GetIDsOfNames(IID_NULL, rgszNames,
			      1+dispParams.cNamedArgs, lcid, rgdispids);

	if (SUCCEEDED(res)) {
	    for (index = 0; index < dispParams.cNamedArgs; ++index) {
		dispParams.rgdispidNamedArgs[index] = rgdispids[index+1];
		res = SetVariantFromSV(hv_iterval(hv, rghe[index]),
				       &dispParams.rgvarg[index], cp);
		if (FAILED(res))
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

	    sv_catpvf(err, " in methodcall/getproperty \"%s\"", buffer);
	}

	for (index = 0; index <= dispParams.cNamedArgs; ++index)
	    SysFreeString(rgszNames[index]);
	Safefree(rgszNames);
	Safefree(rgdispids);

	if (FAILED(res))
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
	    res = SetVariantFromSV(ST(items-1-(index-dispParams.cNamedArgs)),
				   &dispParams.rgvarg[index], cp);
	    if (FAILED(res))
		goto Cleanup;
	}
    }

    res = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid,
				  DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				  &dispParams, &result, &excepinfo, &argErr);

    if (FAILED(res)) {
	/* mega kludge. if a method in WORD is called and we ask
	 * for a result when one is not returned then
	 * hResult == DISP_E_EXCEPTION. this only happens on
	 * functions whose DISPID > 0x8000 */

	if (res == DISP_E_EXCEPTION && dispID > 0x8000) {
	    Zero(&excepinfo, 1, EXCEPINFO);
	    res = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid,
				  DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				  &dispParams, NULL, &excepinfo, &argErr);
	}
    }

    if (SUCCEEDED(res)) {
	if (sv_isobject(retval) && sv_derived_from(retval, szWINOLEVARIANT)) {
	    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(retval);

	    if (pVarObj != NULL) {
		VariantClear(&pVarObj->variant);
		VariantClear(&pVarObj->byref);
		VariantCopy(&pVarObj->variant, &result);
		ST(0) = &sv_yes;
	    }
	}
	else {
	    res = SetSVFromVariant(&result, retval, pObj->stash);
	    ST(0) = &sv_yes;
	}
	VariantClear(&result);
    }
    else {
	err = sv_newmortal();
	sv_setpvf(err, "in methodcall/getproperty \"%s\"", buffer);
	if (res == DISP_E_TYPEMISMATCH || res == DISP_E_PARAMNOTFOUND) {
	    if (argErr < dispParams.cNamedArgs)
		sv_catpvf(err, " argument \"%s\"", hv_iterkey(rghe[argErr], &len));
	    else
		sv_catpvf(err, " argument %d", 1 + dispParams.cArgs - argErr);
	}
    }

 Cleanup:
    if (dispParams.cArgs != 0 && dispParams.rgvarg != NULL) {
	for(index = 0; index < dispParams.cArgs; ++index)
	    VariantClear(&dispParams.rgvarg[index]);
	Safefree(dispParams.rgvarg);
    }
    Safefree(rghe);
    Safefree(dispParams.rgdispidNamedArgs);

    CheckOleError(pObj->stash, res, &excepinfo, err);

    XSRETURN(1);
}

void
GetActiveObject(...)
PPCODE:
{
    CLSID CLSIDObj;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    unsigned int length;
    char *buffer;
    HRESULT res;
    IUnknown *pUnknown;
    IDispatch *pDispatch;

    if (CallObjectMethod(mark, ax, items, "GetActiveObject"))
	return;

    if (items != 2) {
	warn("Usage: Win32::OLE->GetActiveObject(oleclass)");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *oleclass = ST(1);
    UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

    if (!SvPOK(self)) {
	warn("Win32::OLE->GetActiveObject: Must be called as a class method");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    buffer = SvPV(oleclass, length);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ, cp);
    res = CLSIDFromProgID(pBuffer, &CLSIDObj);
    ReleaseBuffer(pBuffer, Buffer);
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    res = GetActiveObject(CLSIDObj, 0, &pUnknown);
    /* Don't call CheckOleError! Return "undef" for "Server not running" */
    if (FAILED(res))
	XSRETURN_UNDEF;

    res = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(stash, pDispatch, NULL);
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
    HRESULT res;

    if (CallObjectMethod(mark, ax, items, "GetObject"))
	return;

    if (items != 2) {
	warn("Usage: Win32::OLE->GetObject(pathname)");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    SV *self = ST(0);
    HV *stash = gv_stashsv(self, TRUE);
    SV *pathname = ST(1);
    UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

    if (!SvPOK(self)) {
	warn("Win32::OLE->GetObject: Must be called as a class method");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    res = CreateBindCtx(0, &pBindCtx);
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    buffer = SvPV(pathname, na);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ, cp);
    res = MkParseDisplayName(pBindCtx, pBuffer, &ulEaten, &pMoniker);
    ReleaseBuffer(pBuffer, Buffer);
    if (FAILED(res)) {
	pBindCtx->Release();
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "after character %lu in \"%s\"", ulEaten, buffer);
	ReportOleError(stash, res, NULL, sv);
	XSRETURN_UNDEF;
    }

    res = pMoniker->BindToObject(pBindCtx, NULL, IID_IDispatch,
				 (void**)&pDispatch);
    pBindCtx->Release();
    pMoniker->Release();
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(stash, pDispatch, NULL);
    XSRETURN(1);
}

void
QueryObjectType(...)
PPCODE:
{
    if (CallObjectMethod(mark, ax, items, "QueryObjectType"))
	return;

    if (items != 2) {
	warn("Usage: Win32::OLE->QueryObjectType(object)");
	DEBUGBREAK;
	XSRETURN_UNDEF;
    }

    SV *object = ST(1);

    if (!sv_isobject(object) || !sv_derived_from(object, szWINOLE))
	XSRETURN_UNDEF;

    WINOLEOBJECT *pObj = GetOleObject(object);
    if (pObj == NULL)
	XSRETURN_UNDEF;

    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int count;
    BSTR bstr;
    char szName[64];
    char *pszName;

    HRESULT res = pObj->pDispatch->GetTypeInfoCount(&count);
    if (FAILED(res) || count == 0)
	XSRETURN_UNDEF;

    HV *stash = gv_stashsv(ST(0), TRUE);
    LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

    res = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    /* Return ('TypeLib Name', 'Class Name') in array context */
    if (GIMME_V == G_ARRAY) {
	res = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	if (FAILED(res)) {
	    pTypeInfo->Release();
	    ReportOleError(stash, res, NULL, NULL);
	    XSRETURN_UNDEF;
	}

	res = pTypeLib->GetDocumentation(-1, &bstr, NULL, NULL, NULL);
	pTypeLib->Release();
	if (FAILED(res)) {
	    pTypeInfo->Release();
	    ReportOleError(stash, res, NULL, NULL);
	    XSRETURN_UNDEF;
	}

	pszName = GetMultiByte(bstr, szName, sizeof(szName), cp);
	PUSHs(sv_2mortal(newSVpv(pszName, 0)));
	SysFreeString(bstr);
	ReleaseBuffer(pszName, szName);
    }

    res = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstr, NULL, NULL, NULL);
    pTypeInfo->Release();
    if (CheckOleError(stash, res, NULL, NULL))
	XSRETURN_UNDEF;

    pszName = GetMultiByte(bstr, szName, sizeof(szName), cp);
    PUSHs(sv_2mortal(newSVpv(pszName, 0)));
    SysFreeString(bstr);
    ReleaseBuffer(pszName, szName);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Tie

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(self);
    if (pObj != NULL) {
	DBG(("Win32::OLE::Tie::DESTROY |%lx| |%lx|\n", pObj, pObj->pDispatch));
	DestroyPerlObject(pObj);
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
    SV **coo;
    char *buffer;
    unsigned int length;
    unsigned int argErr;
    WINOLEOBJECT *pObj;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    VARIANTARG propName;
    DISPID dispID = DISPID_VALUE;
    HRESULT res;

    ST(0) = &sv_undef;

    coo = hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DBG(("Win32::OLE::Tie::FETCH |%s| |%d| |%lx|\n",
	 PERL_OLE_ID, PERL_OLE_IDLEN, coo));

    if (coo == NULL) {
	warn("Win32::OLE::Tie::FETCH: Not a Win32::OLE object");
	DEBUGBREAK;
	XSRETURN(1);
    }

    buffer = SvPV(key, length);
    if (strEQ(buffer, PERL_OLE_ID)) {
	ST(0) = *coo;
	XSRETURN(1);
    }

    pObj = CheckOleStruct(SvIV(*coo));
    if (pObj == NULL) {
	XSRETURN(1);
    }

    VariantInit(&result);
    VariantInit(&propName);

    LCID lcid = QueryPkgVar(pObj->stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(pObj->stash, CP_NAME, CP_LEN, cpDefault);

    dispParams.cArgs = 0;
    dispParams.rgvarg = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    res = GetHashedDispID(pObj, buffer, length, dispID, lcid, cp);
    if (FAILED(res)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(pObj->stash, res, NULL, sv_2mortal(err));
	    XSRETURN(1);
	}

	/* default method call: $self->{Key} ---> $self->Item('Key') */
	V_VT(&propName) = VT_BSTR;
	V_BSTR(&propName) = AllocOleString(buffer, length, cp);
	dispParams.cArgs = 1;
	dispParams.rgvarg = &propName;
    }

    Zero(&excepinfo, 1, EXCEPINFO);

    res = pObj->pDispatch->Invoke(dispID, IID_NULL,
		    lcid, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
		    &dispParams, &result, &excepinfo, &argErr);

    VariantClear(&propName);

    if (FAILED(res)) {
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "in methodcall/getproperty \"%s\"", buffer);
	ReportOleError(pObj->stash, res, &excepinfo, sv);
    }
    else {
	ST(0) = sv_newmortal();
	res = SetSVFromVariant(&result, ST(0), pObj->stash);
	CheckOleError(pObj->stash, res, NULL, NULL);
    }
    VariantClear(&result);

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
    HRESULT res;
    EXCEPINFO excepinfo;
    DISPID dispID = DISPID_VALUE;
    DISPID dispIDParam;
    DISPPARAMS dispParams;
    VARIANTARG propertyValue[2];

    WINOLEOBJECT *pObj = GetOleObject(self);
    if (pObj == NULL)
	XSRETURN_EMPTY;

    LCID lcid = QueryPkgVar(pObj->stash, LCID_NAME, LCID_LEN, lcidDefault);
    UINT cp = QueryPkgVar(pObj->stash, CP_NAME, CP_LEN, cpDefault);

    dispIDParam = DISPID_PROPERTYPUT;
    dispParams.rgdispidNamedArgs = &dispIDParam;
    dispParams.rgvarg = propertyValue;
    dispParams.cNamedArgs = 1;
    dispParams.cArgs = 1;

    VariantInit(&propertyValue[0]);
    VariantInit(&propertyValue[1]);

    buffer = SvPV(key, length);
    res = GetHashedDispID(pObj, buffer, length, dispID, lcid, cp);
    if (FAILED(res)) {
	if (!SvTRUE(def)) {
	    SV *err = newSVpvf(" in GetIDsOfNames \"%s\"", buffer);
	    ReportOleError(pObj->stash, res, NULL, sv_2mortal(err));
	    XSRETURN_EMPTY;
	}

	dispParams.cArgs = 2;
	V_VT(&propertyValue[1]) = VT_BSTR;
	V_BSTR(&propertyValue[1]) = AllocOleString(buffer, length, cp);
    }

    res = SetVariantFromSV(value, &propertyValue[0], cp);
    if (CheckOleError(pObj->stash, res, NULL, NULL)) {
	VariantClear(&propertyValue[1]);
	XSRETURN_EMPTY;
    }

    Zero(&excepinfo, 1, EXCEPINFO);
    res = pObj->pDispatch->Invoke(dispID, IID_NULL, lcid, DISPATCH_PROPERTYPUT,
				  &dispParams, NULL, &excepinfo, &argErr);

    for(index = 0; index < dispParams.cArgs; ++index)
	VariantClear(&propertyValue[index]);

    if (FAILED(res)) {
	SV *err = sv_newmortal();
	sv_setpvf(err, "in setproperty \"%s\"", buffer);
	ReportOleError(pObj->stash, res, &excepinfo, err);
    }

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
    WINOLEOBJECT *pObj = GetOleObject(self);
    DBG(("FIRST/NEXTKEY (%d) called, pObj=%p\n", ix, pObj));
    if (pObj == NULL)
	XSRETURN_UNDEF;

    switch (ix)
    {
    case 0: /* FIRSTKEY */
	FetchTypeInfo(pObj);
	pObj->PropIndex = 0;
    case 1: /* NEXTKEY */
	ST(0) = NextPropertyName(pObj);
	break;

    case 2: /* FIRSTENUM */
	if (pObj->pEnum != NULL)
	    pObj->pEnum->Release();
	pObj->pEnum = CreateEnumVARIANT(pObj);
    case 3: /* NEXTENUM */
	ST(0) = NextEnumElement(pObj->pEnum, pObj->stash);
	if (!SvOK(ST(0))) {
	    pObj->pEnum->Release();
	    pObj->pEnum = NULL;
	}
	break;
    }

    if (!SvREADONLY(ST(0)))
	sv_2mortal(ST(0));

    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Const

void
_Load(clsid,major,minor,locale,codepage,typelib)
    SV *clsid
    IV major
    IV minor
    SV *locale
    SV *codepage
    SV *typelib
PPCODE:
{
    ITypeLib *pTypeLib;
    CLSID CLSIDObj;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    HRESULT res;
    LCID lcid = lcidDefault;
    UINT cp = cpDefault;
    HV *stash = gv_stashpv(szWINOLE, TRUE);

    if (SvIOK(locale))
	lcid = SvIV(locale);

    if (SvIOK(codepage))
	cp = SvIV(codepage);

    if (sv_derived_from(clsid, szWINOLE)) {
	/* Get containing typelib from IDispatch interface */
	ITypeInfo *pTypeInfo;
	unsigned int count;
	WINOLEOBJECT *pObj = GetOleObject(clsid);
	if (pObj == NULL)
	    XSRETURN_UNDEF;

	stash = pObj->stash;
	res = pObj->pDispatch->GetTypeInfoCount(&count);
	if (CheckOleError(stash, res, NULL, NULL) || count == 0)
	    XSRETURN_UNDEF;

	lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);
	cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

	res = pObj->pDispatch->GetTypeInfo(0, lcid, &pTypeInfo);
	if (CheckOleError(stash, res, NULL, NULL))
	    XSRETURN_UNDEF;

	res = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	pTypeInfo->Release();
	if (CheckOleError(stash, res, NULL, NULL))
	    XSRETURN_UNDEF;
    }
    else {
	/* try to load registered typelib by clsid, version and lcid */
	char *pszBuffer = SvPV(clsid, na);
	pBuffer = GetWideChar(pszBuffer, Buffer, OLE_BUF_SIZ, cp);
	res = CLSIDFromString(pBuffer, &CLSIDObj);
	ReleaseBuffer(pBuffer, Buffer);

	if (CheckOleError(stash, res, NULL, NULL))
	    XSRETURN_UNDEF;

	res = LoadRegTypeLib(CLSIDObj, major, minor, lcid, &pTypeLib);
	if (FAILED(res) && SvPOK(typelib)) {
	    /* typelib not registerd, try to read from file "typelib" */
	    pszBuffer = SvPV(typelib, na);
	    pBuffer = GetWideChar(pszBuffer, Buffer, OLE_BUF_SIZ, cp);
	    res = LoadTypeLib(pBuffer, &pTypeLib);
	    ReleaseBuffer(pBuffer, Buffer);
	}
	if (CheckOleError(stash, res, NULL, NULL))
	    XSRETURN_UNDEF;
    }

    /* we'll return ref to hash with constant name => value pairs */
    HV *hv = newHV();
    unsigned int count = pTypeLib->GetTypeInfoCount();

    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));

    /* loop through all objects in type lib */
    for (int index=0 ; index < count ; ++index) {
	ITypeInfo *pTypeInfo;
	LPTYPEATTR pTypeAttr;

	res = pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (CheckOleError(stash, res, NULL, NULL))
	    continue;

	res = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (FAILED(res)) {
	    pTypeInfo->Release();
	    ReportOleError(stash, res, NULL, NULL);
	    continue;
	}

	/* extract all constants for each ENUM */
	if (pTypeAttr->typekind == TKIND_ENUM) {
	    for (int iVar=0 ; iVar < pTypeAttr->cVars ; ++iVar) {
		LPVARDESC pVarDesc;

		res = pTypeInfo->GetVarDesc(iVar, &pVarDesc);
		/* XXX LEAK alert */
		if (CheckOleError(stash, res, NULL, NULL))
		    continue;

		if (pVarDesc->varkind == VAR_CONST &&
		    !(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
					     VARFLAG_FRESTRICTED |
					     VARFLAG_FNONBROWSABLE))) {
		    unsigned int cName;
		    BSTR bstr;
		    char szName[64];

		    res = pTypeInfo->GetNames(pVarDesc->memid, &bstr,
					      1, &cName);
		    if (CheckOleError(stash, res, NULL, NULL)
			|| cName == 0 || bstr == NULL)
			continue;

		    char *pszName = GetMultiByte(bstr, szName, sizeof(szName),
						 cp);
		    SV *sv = newSVpv("",0);
		    /* XXX LEAK alert */
		    res = SetSVFromVariant(pVarDesc->lpvarValue, sv, stash);
		    if (!CheckOleError(stash, res, NULL, NULL))
			hv_store(hv, pszName, strlen(pszName), sv, 0);

		    SysFreeString(bstr);
		    ReleaseBuffer(pszName, szName);
		}
		pTypeInfo->ReleaseVarDesc(pVarDesc);
	    }
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
	WINOLEOBJECT *pObj = GetOleObject(object);
	if (pObj != NULL) {
	    pEnumObj->pEnum = CreateEnumVARIANT(pObj);
	    pEnumObj->stash = pObj->stash;
	}
    }
    else { /* Clone */
	WINOLEENUMOBJECT *pOriginal = GetOleEnumObject(self);
	HRESULT res = pOriginal->pEnum->Clone(&pEnumObj->pEnum);
	CheckOleError(pOriginal->stash, res, NULL, NULL);
	pEnumObj->stash = pOriginal->stash;
    }

    if (pEnumObj->pEnum == NULL) {
	Safefree(pEnumObj);
	XSRETURN_UNDEF;
    }

    AddToObjectChain((OBJECTHEADER*)pEnumObj, WINOLEENUM_MAGIC);
    SvREFCNT_inc(pEnumObj->stash);

    SV *sv = newSViv((IV)pEnumObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), GetStash(self)));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(self);
    RemoveFromObjectChain((OBJECTHEADER*)pEnumObj);
    HRESULT res = pEnumObj->pEnum->Release();
    CheckOleError(pEnumObj->stash, res, NULL, NULL);
    XSRETURN_EMPTY;
}

void
Next(self,...)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(self);
    IV count = 1;
    if (items > 1)
	count = SvIV(ST(1));

    if (count < 1) {
	warn("Win32::OLE::Enum::Next: invalid Count %ld", count);
	DEBUGBREAK;
	count = 1;
    }

    SV *sv = NULL;
    while (count-- > 0) {
	sv = NextEnumElement(pEnumObj->pEnum, pEnumObj->stash);
	if (!SvOK(sv))
	    break;
	if (!SvREADONLY(sv))
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
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(self);
    HRESULT res = pEnumObj->pEnum->Reset();
    CheckOleError(pEnumObj->stash, res, NULL, NULL);
    ST(0) = (res == S_OK) ? &sv_yes : &sv_no;
    XSRETURN(1);
}

void
Skip(self,...)
    SV *self
PPCODE:
{
    WINOLEENUMOBJECT *pEnumObj = GetOleEnumObject(self);
    IV count = 1;
    if (items > 1)
	count = SvIV(ST(1));
    HRESULT res = pEnumObj->pEnum->Skip(count);
    CheckOleError(pEnumObj->stash, res, NULL, NULL);
    ST(0) = (res == S_OK) ? &sv_yes : &sv_no;
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
    HV *stash = GetStash(self);
    HRESULT res;
    WINOLEVARIANTOBJECT *pVarObj;

    New(0, pVarObj, 1, WINOLEVARIANTOBJECT);
    VariantInit(&pVarObj->variant);
    VariantInit(&pVarObj->byref);
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
	LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);
	UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

	V_VT(&pVarObj->variant) = VT_BSTR;
	ptr = SvPV(data, length);
	V_BSTR(&pVarObj->variant) = AllocOleString(ptr, length, cp);
	VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant, lcid,0, vt);
	break;
    }

    case VT_BSTR:
    {
	UINT cp = QueryPkgVar(stash, CP_NAME, CP_LEN, cpDefault);

	ptr = SvPV(data, length);
	V_BSTR(&pVarObj->variant) = AllocOleString(ptr, length, cp);
	break;
    }

    case VT_DISPATCH:
    {
	/* Argument MUST be a valid Perl OLE object! */
	WINOLEOBJECT *pObj = GetOleObject(data);
	if (pObj == NULL)
	    V_VT(&pVarObj->variant) = VT_EMPTY;
	else {
	    pObj->pDispatch->AddRef();
	    V_DISPATCH(&pVarObj->variant) = pObj->pDispatch;
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
	WINOLEOBJECT *pObj = GetOleObject(data);
	if (pObj == NULL)
	    V_VT(&pVarObj->variant) = VT_EMPTY;
	else {
	    res = pObj->pDispatch->QueryInterface(IID_IUnknown,
			           (void**)&V_UNKNOWN(&pVarObj->variant));
	    CheckOleError(pObj->stash, res, NULL, NULL);
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
		res = SafeArrayAccessData(V_ARRAY(&pVarObj->variant),
						   (void**)&pDest);
		if (FAILED(res)) {
		    VariantClear(&pVarObj->variant);
		    ReportOleError(stash, res, NULL, NULL);
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
	warn("Win32::OLE::Variant::new: Invalid value type %d",
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

    AddToObjectChain((OBJECTHEADER*)pVarObj, WINOLEVARIANT_MAGIC);

    SV *sv = newSViv((IV)pVarObj);
    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(self);
    if (pVarObj != NULL) {
	RemoveFromObjectChain((OBJECTHEADER*)pVarObj);
	VariantClear(&pVarObj->byref);
	VariantClear(&pVarObj->variant);
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
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(self);

    ST(0) = &sv_undef;
    if (pVarObj != NULL) {
	ST(0) = sv_newmortal();
	if (ix == 0)
	    sv_setiv(ST(0), V_VT(&pVarObj->variant));
	else
	    SetSVFromVariant(&pVarObj->variant, ST(0), SvSTASH(SvRV(self)));
    }
    XSRETURN(1);
}

void
As(self,type)
    SV *self
    IV type
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(self);

    ST(0) = &sv_undef;
    if (pVarObj != NULL) {
	HRESULT res;
	VARIANT variant;
	HV *stash = GetStash(self);
	LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);

	VariantInit(&variant);
	res = VariantChangeTypeEx(&variant, &pVarObj->variant, lcid, 0, type);
	if (!CheckOleError(stash, res, NULL, NULL)) {
	    ST(0) = sv_newmortal();
	    SetSVFromVariant(&variant, ST(0), SvSTASH(SvRV(self)));
	}
    }
    XSRETURN(1);
}

void
ChangeType(self,type)
    SV *self
    IV type
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(self);
    HRESULT res = E_INVALIDARG;

    if (pVarObj != NULL) {
	HV *stash = GetStash(self);
	LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);

	res = VariantChangeTypeEx(&pVarObj->variant, &pVarObj->variant, 
				  lcid, 0, type);
	CheckOleError(stash, res, NULL, NULL);
    }

    if (FAILED(res))
	ST(0) = &sv_undef;

    XSRETURN(1);
}

void
Unicode(self)
    SV *self
PPCODE:
{
    WINOLEVARIANTOBJECT *pVarObj = GetOleVariantObject(self);

    ST(0) = &sv_undef;
    if (pVarObj != NULL) {
	HV *stash = GetStash(self);
	VARIANT Variant;
	VARIANT *pVariant = &pVarObj->variant;
	HRESULT res = S_OK;

	if ((V_VT(pVariant) & ~VT_BYREF) != VT_BSTR) {
	    LCID lcid = QueryPkgVar(stash, LCID_NAME, LCID_LEN, lcidDefault);

	    VariantInit(&Variant);
	    res = VariantChangeTypeEx(&Variant, pVariant, lcid, 0, VT_BSTR);
	    pVariant = &Variant;
	}

	if (!CheckOleError(stash, res, NULL, NULL)) {
	    BSTR bstr = V_ISBYREF(pVariant) ? *V_BSTRREF(pVariant) 
		                            : V_BSTR(pVariant);
	    STRLEN len = SysStringLen(bstr);
	    SV *sv = newSVpv((char*)bstr, 2*len);
	    U16 *pus = (U16 *)SvPV(sv, na);
	    for (STRLEN i=0 ; i < len ; ++i)
		pus[i] = htons(pus[i]);

	    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), 
					gv_stashpv("Unicode::String", TRUE)));
	}
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
	SvCUR(sv) = LCMapStringA(lcid, flags, string, length, SvPVX(sv), SvLEN(sv));
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
