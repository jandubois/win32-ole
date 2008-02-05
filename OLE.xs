/* OLE.xs
 *
 *  (c) 1995 Microsoft Corporation. All rights reserved.
 *  Developed by ActiveWare Internet Corp., http://www.ActiveWare.com
 *
 *  Other modifications (c) 1997 by Gurusamy Sarathy <gsar@umich.edu>
 *  and Jan Dubois <jan.dubois@ibm.net>
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
 * - Package Win32::OLE::Enum     OLE collection enumeratation
 * - Package Win32::OLE::Variant  Implements Perl VARIANT objects
 *
 */

#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#include <windows.h>

#if defined(__cplusplus)
extern "C" {
#endif

#include "EXTERN.h"
#include "perl.h"
#include "XSub.h"

#undef DEB

// #define MYDEBUG

#if !defined(MYDEBUG)
#    define DEB(a)
#else
#    define DEB(a)  MyDebug a;
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
static const int OLE_BUF_SIZ = 256;
static const LCID lcidDefault = (0x02 << 10) /* LOCALE_SYSTEM_DEFAULT */;
static char PERL_OLE_ID[] = "___Perl___OleObject___";
static const int PERL_OLE_IDLEN = sizeof(PERL_OLE_ID)-1;

/* Perl OLE object definition */
typedef struct _tagWINOLEOBJECT WINOLEOBJECT;
typedef struct _tagWINOLEOBJECT
{
    long Win32OleMagic;

    WINOLEOBJECT *pNext;
    WINOLEOBJECT *pPrevious;

    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo;

    HV *stash;
    HV *hashTable;
    SV *destroy;

    unsigned short cFuncs;
    unsigned short cVars;
    unsigned int   PropIndex;

}   WINOLEOBJECT;

/* global variables */
static WINOLEOBJECT *g_pObj = NULL;
static long LastOleError;


/* Calling Carp::croak gives a better error location than croak alone. */
void
DoCroak(const char *pat, ...)
{
    dSP;
    SV *sv;

    va_list args;
    va_start(args, pat);
    sv_vsetpvfn(sv, pat, strlen(pat), &args, Null(SV**), 0, Null(bool*));
    va_end(args);

    PUSHMARK(sp) ;
    XPUSHs(sv);
    PUTBACK;
    perl_call_pv("Carp::croak", G_DISCARD);
    /* NOTREACHED */
}

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
GetMultiByte(OLECHAR *wide, char *psz, int len)
{
    int count = WideCharToMultiByte(CP_ACP, 0, wide, -1, psz, len, NULL, NULL);
    if (count > 0)
	return psz;

    count = WideCharToMultiByte(CP_ACP, 0, wide, -1, NULL, 0, NULL, NULL);
    if (count == 0) /* should never happen :-) */
	DoCroak("Win32::OLE [GetMultiByte] failure: %lu", GetLastError());

    New(4711, psz, count, char);
    count = WideCharToMultiByte(CP_ACP, 0, wide, -1, psz, count, NULL, NULL);
    return psz;
}

OLECHAR *
GetWideChar(char *psz, OLECHAR *wide, int len)
{
    /* Note: len is number of OLECHARs, not bytes! */
    int count = MultiByteToWideChar(CP_ACP, 0, psz, -1, wide, len);
    if (count > 0)
	return wide;

    count = MultiByteToWideChar(CP_ACP, 0, psz, -1, NULL, 0);
    if (count == 0)
	DoCroak("Win32::OLE [GetWideChar] failure: %lu", GetLastError());

    New(4711, wide, count, OLECHAR);
    count = MultiByteToWideChar(CP_ACP, 0, psz, -1, wide, count);
    return wide;
}

void
ReleaseExcepInfo(EXCEPINFO *pExcepInfo)
{
    if (pExcepInfo != NULL) {
	SysFreeString(pExcepInfo->bstrSource);
	SysFreeString(pExcepInfo->bstrDescription);
	SysFreeString(pExcepInfo->bstrHelpFile);
    }
}

void
ReportOleError(HRESULT LastError, EXCEPINFO *pExcepInfo, SV *svDetails)
{
    dSP;

    if (!dowarn) {
	ReleaseExcepInfo(pExcepInfo);
	return;
    }

    SV *sv = sv_2mortal(newSVpv("", 0));

    /* start with exception info */
    if (pExcepInfo != NULL && (pExcepInfo->bstrSource != NULL ||
			       pExcepInfo->bstrDescription != NULL )) {
	char szSource[80] = "<Unknown Source>";
	char szDescription[200] = "<No description provided>";

	char *pszSource = szSource;
	char *pszDescription = szDescription;

	if (pExcepInfo->bstrSource != NULL)
	    pszSource = GetMultiByte(pExcepInfo->bstrSource,
				     szSource, sizeof(szSource));
	
	if (pExcepInfo->bstrDescription != NULL)
	    pszDescription = GetMultiByte(pExcepInfo->bstrDescription,
				     szDescription, sizeof(szDescription));
	
	sv_setpvf(sv, "OLE exception from \"%s\":\n\n%s\n\n",
		  pszSource, pszDescription);

	ReleaseBuffer(pszSource, szSource);
	ReleaseBuffer(pszDescription, szDescription);
	ReleaseExcepInfo(pExcepInfo);
    }

    /* always include OLE error code */
    sv_catpvf(sv, "OLE error 0x%08x", LastError);

    /* try to append ': "error text"' from message catalog */
    char *pszMsgText;
    DWORD dwCount = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
				  FORMAT_MESSAGE_FROM_SYSTEM |
				  FORMAT_MESSAGE_IGNORE_INSERTS,
				  NULL, LastError, lcidDefault,
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

    PUSHMARK(sp) ;
    XPUSHs(sv);
    PUTBACK;
    perl_call_pv("Carp::croak", G_DISCARD);
    /* NOTREACHED */

}   /* ReportOleError */

inline BOOL
CheckOleError(HRESULT LastError, EXCEPINFO *pExcepInfo, SV *svDetails)
{
    if (FAILED(LastError)) {
	ReportOleError(LastError, pExcepInfo, svDetails);
	return TRUE;
    }
    return FALSE;
}

SV *
CreatePerlObject(HV *stash, IDispatch *pDispatch, SV *destroy)
{
    /* returns a mortal reference to a new Perl OLE object */

    if (pDispatch == NULL)
	DoCroak("Win32::OLE [CreatePerlObject]: Invalid IDispatch interface");

    HV *hvouter = newHV();
    HV *hvinner = newHV();
    SV *inner;
    WINOLEOBJECT *pObj;

    New(2101, pObj, 1, WINOLEOBJECT);
    pObj->Win32OleMagic = WINOLE_MAGIC;
    pObj->pPrevious = NULL;
    pObj->pDispatch = pDispatch;
    pObj->pTypeInfo = NULL;
    pObj->hashTable = newHV();
    pObj->destroy = destroy;
    pObj->stash = stash;
    SvREFCNT_inc(stash);

    pObj->pNext = g_pObj;
    if (g_pObj)
	g_pObj->pPrevious = pObj;
    g_pObj = pObj;

    DEB(("CreatePerlObject = |%lx| Class = %s\n", pObj, HvNAME(stash)));

    hv_store(hvinner, PERL_OLE_ID, PERL_OLE_IDLEN, newSViv((long)pObj), 0);
    inner = sv_bless(newRV_noinc((SV*)hvinner), 
		     gv_stashpv("Win32::OLE::Tie", TRUE));
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
    if (pObj->pDispatch == NULL)
	DoCroak("Win32::OLE [DestroyPerlObject]: Object is already destroyed");

    DEB(("DestroyPerlObject |%lx|", pObj));

    DEB((" pDispatch"));
    pObj->pDispatch->Release();
    pObj->pDispatch = NULL;

    DEB((" stash(%d)", SvREFCNT(pObj->stash)));
    SvREFCNT_inc(pObj->stash);

    DEB((" hashTable(%d)", SvREFCNT(pObj->hashTable)));
    SvREFCNT_dec(pObj->hashTable);
    pObj->hashTable = NULL;

    if (pObj->pTypeInfo != NULL) {
	DEB((" pTypeInfo"));
	pObj->pTypeInfo->Release();
	pObj->pTypeInfo = NULL;
    }

    if (pObj->destroy != NULL) {
	DEB((" destroy(%d)", SvREFCNT(pObj->destroy)));
	SvREFCNT_dec(pObj->destroy);
	pObj->destroy = NULL;
    }

    DEB(("\n"));
    Safefree(pObj);

}   /* DestroyPerlObject */

WINOLEOBJECT *
CheckOleStruct(IV addr)
{
    WINOLEOBJECT *pObj = (WINOLEOBJECT*)addr;

    if (pObj == NULL || pObj->Win32OleMagic != WINOLE_MAGIC)
	DoCroak("Win32::OLE [CheckOleStruct]: Damaged Win32::OLE object");

    return pObj;
}

WINOLEOBJECT *
GetOleObject(SV *sv)
{
    if (sv != NULL && SvROK(sv)) {
	SV **psv = hv_fetch((HV*)SvRV(sv), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
	if (psv != NULL) {
	    DEB(("GetOleObject = |%lx|\n", SvIV(*psv)));
	    return CheckOleStruct(SvIV(*psv));
	}
    }
    DoCroak("Win32::OLE [GetOleObject]: Damaged Win32::OLE object");
    return (WINOLEOBJECT*)NULL;
}

BSTR
AllocOleString(char* pStr, int length)
{
    int count = MultiByteToWideChar(CP_ACP, 0, pStr, length, NULL, 0);
    BSTR bstr = SysAllocStringLen(NULL, count);
    MultiByteToWideChar(CP_ACP, 0, pStr, length, bstr, count);
    return bstr;
}

BOOL
GetHashedDispID(WINOLEOBJECT *pObj, char *buffer, STRLEN len, DISPID &dispID)
{
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

    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ);
    LastOleError = pObj->pDispatch->
	GetIDsOfNames(IID_NULL, &pBuffer, 1, lcidDefault, &id);
    ReleaseBuffer(pBuffer, Buffer);
    /* Don't call CheckOleError! Caller might retry the "unnamed" method */
    if (FAILED(LastOleError))
	return FALSE;

    hv_store(pObj->hashTable, buffer, len, newSViv(id), 0);

    dispID = id;
    return TRUE;

}   /* GetHashedDispID */

void
FetchTypeInfo(WINOLEOBJECT *pObj)
{
    unsigned int count;
    LPTYPEATTR pTypeAttr;

    if (pObj->pTypeInfo != NULL)
	return;

    if (FAILED(pObj->pDispatch->GetTypeInfoCount(&count)))
	DoCroak("Win32::OLE [FetchTypeInfo]: GetTypeInfoCount failed\n");

    if (count == 0) {
	warn("Win32::OLE [FetchTypeInfo]: GetTypeInfoCount returned 0");
	return;
    }

    LastOleError = pObj->pDispatch->
	GetTypeInfo(0, lcidDefault, &pObj->pTypeInfo);
    if (CheckOleError(LastOleError, NULL, NULL))
	return;

    LastOleError = pObj->pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (CheckOleError(LastOleError, NULL, NULL)) {
	pObj->pTypeInfo->Release();
	pObj->pTypeInfo = NULL;
	return;
    }

    pObj->cFuncs = pTypeAttr->cFuncs;
    pObj->cVars = pTypeAttr->cVars;
    pObj->PropIndex = 0;
    pObj->pTypeInfo->ReleaseTypeAttr(pTypeAttr);

}   /* FetchTypeInfo */

SV *
NextPropertyName(WINOLEOBJECT *pObj)
{
    unsigned int cName;
    BSTR bstr;
    char szName[64];

    while (pObj->PropIndex < pObj->cFuncs+pObj->cVars) {
	ULONG index = pObj->PropIndex++;
	/* Try all the INVOKE_PROPERTYGET functions first */
	if (index < pObj->cFuncs) {
	    LPFUNCDESC pFuncDesc;

	    LastOleError = pObj->pTypeInfo->GetFuncDesc(index, &pFuncDesc);
	    if (CheckOleError(LastOleError, NULL, NULL))
		continue;

	    if (!(pFuncDesc->funckind & FUNC_DISPATCH) ||
		!(pFuncDesc->invkind & INVOKE_PROPERTYGET) ||
	        (pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED |
					  FUNCFLAG_FHIDDEN |
					  FUNCFLAG_FNONBROWSABLE))) {
		pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		continue;
	    }

	    LastOleError = pObj->pTypeInfo->GetNames(pFuncDesc->memid,
						      &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseFuncDesc(pFuncDesc);
	    if (CheckOleError(LastOleError, NULL, NULL)
		|| cName == 0 || bstr == NULL)
		continue;

	    char *pszName = GetMultiByte(bstr, szName, sizeof(szName));
	    SV *sv = newSVpv(pszName, 0);
	    SysFreeString(bstr);
	    ReleaseBuffer(pszName, szName);
	    return sv;
	}
	/* Now try the VAR_DISPATCH kind variables used by older OLE versions */
	else {
	    LPVARDESC pVarDesc;

	    index -= pObj->cFuncs;
	    LastOleError = pObj->pTypeInfo->GetVarDesc(index, &pVarDesc);
	    if (CheckOleError(LastOleError, NULL, NULL))
		continue;

	    if (!(pVarDesc->varkind & VAR_DISPATCH) ||
		(pVarDesc->wVarFlags & (VARFLAG_FRESTRICTED |
					VARFLAG_FHIDDEN |
					VARFLAG_FNONBROWSABLE))) {
		pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
		continue;
	    }

	    LastOleError = pObj->pTypeInfo->GetNames(pVarDesc->memid,
						     &bstr, 1, &cName);
	    pObj->pTypeInfo->ReleaseVarDesc(pVarDesc);
	    if (CheckOleError(LastOleError, NULL, NULL)
		|| cName == 0 || bstr == NULL)
		continue;

	    char *pszName = GetMultiByte(bstr, szName, sizeof(szName));
	    SV *sv = newSVpv(pszName, 0);
	    SysFreeString(bstr);
	    ReleaseBuffer(pszName, szName);
	    return sv;
	}
    }
    return &sv_undef;

}   /* NextPropertyName */

void
SetVariantFromSV(SV* sv, VARIANT *pVariant)
{
    VariantInit(pVariant);

    /* XXX requirement to call mg_get() may change in Perl > 5.004 */
    if (SvGMAGICAL(sv))
	mg_get(sv);

    /* Objects */
    if (SvROK(sv)) {
	if (sv_derived_from(sv, "Win32::OLE")) {
	    IDispatch *pDispatch = GetOleObject(sv)->pDispatch;
	    pDispatch->AddRef();
	    V_VT(pVariant) = VT_DISPATCH;
	    V_DISPATCH(pVariant) = pDispatch;
	    return;
	}

	if (sv_derived_from(sv, "Win32::OLE::Variant")) {
	    STRLEN len;
	    VARIANT *pPerlVariant = (VARIANT*)SvPV(SvRV(sv), len);

	    if (len != sizeof(VARIANT))
		DoCroak("Win32::OLE [SetVariantFromSV]: Invalid object");

	    LastOleError = VariantCopy(pVariant, pPerlVariant);
	    CheckOleError(LastOleError, NULL, NULL);
	    return;
	}

	sv = SvRV(sv);
    }

    /* Arrays */
    if (SvTYPE(sv) == SVt_PVAV) {
	AV *av = (AV*)sv;
	IV len = av_len(av)+1;
	VARIANT variant;

	V_ARRAY(pVariant) = SafeArrayCreateVector(VT_VARIANT, 0, len);
	if (V_ARRAY(pVariant) == NULL) {
	    CheckOleError(E_OUTOFMEMORY, NULL, NULL);
	    return;
	}

	V_VT(pVariant) = VT_VARIANT | VT_ARRAY;
	for (IV index=0; index < len ; ++index) {
	    SV **psv = av_fetch(av, index, 0);
	    if (psv != NULL) {
		SetVariantFromSV(*psv, &variant);
		LastOleError = SafeArrayPutElement(V_ARRAY(pVariant),
						   &index, &variant);
		CheckOleError(LastOleError, NULL, NULL);
	    }
	}
	return;
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
	V_BSTR(pVariant) = AllocOleString(ptr, len);
    }
    else {
	V_VT(pVariant) = VT_ERROR;
	V_ERROR(pVariant) = DISP_E_PARAMNOTFOUND;
    }
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

SV *
SetSVFromVariant(VARIANTARG *pVariant, SV* sv, HV *stash)
{
    sv_setsv(sv, &sv_undef);

    if (V_ISARRAY(pVariant)) {
	AV *av;
	VARIANT variant;
	int dim, index;
	long *pArrayIndex, *pLowerBound, *pUpperBound;
	HRESULT hResult;

	dim = SafeArrayGetDim(V_ARRAY(pVariant));
	New(4444, pArrayIndex, dim, long);
	New(4444, pLowerBound, dim, long);
	New(4444, pUpperBound, dim, long);
	for(index = 1; index <= dim; ++index) {
	    hResult = SafeArrayGetLBound(V_ARRAY(pVariant), index,
					  &pLowerBound[index-1]);
	    if (FAILED(hResult))
		goto ErrorExit;
	}

	for(index = 1; index <= dim; ++index) {
	    hResult = SafeArrayGetUBound(V_ARRAY(pVariant), index,
					  &pUpperBound[index-1]);
	    if (FAILED(hResult))
		goto ErrorExit;
	}

	av = newAV();
	if (dim < 3)
	{
	    memcpy(pArrayIndex, pLowerBound, dim*sizeof(long));
	    for(index = dim-1;
		pArrayIndex[index] <= pUpperBound[index];
		++pArrayIndex[index])
	    {
		hResult = SafeArrayGetElement(V_ARRAY(pVariant), pArrayIndex,
					      &variant);
		if (SUCCEEDED(hResult)) {
		    av_push(av, SetSVFromVariant(&variant, newSVpv("",0),
						 stash));
		}
	    }
	}
	sv = newRV_noinc((SV*)av);

ErrorExit:
	Safefree(pArrayIndex);
	Safefree(pLowerBound);
	Safefree(pUpperBound);
	return sv;
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

	if (V_ISBYREF(pVariant))
	    pStr = GetMultiByte(*V_BSTRREF(pVariant), Str, sizeof(Str));
	else
	    pStr = GetMultiByte(V_BSTR(pVariant), Str, sizeof(Str));

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
	HRESULT hResult = VariantChangeType(pVariant, pVariant,
					    0, VT_BSTR);
	if (SUCCEEDED(hResult))
	    goto ConvertString;
	break;
    }
    }

    return sv;

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
	OleInitialize(NULL);
	break;

    case DLL_PROCESS_DETACH:
	/* Global destruction will have normally DESTROYed all
	 * objects, so the loop below will never be entered.
	 * Unless global destruction phase was somehow interrupted.
	 * Only external resources are cleaned up here.
	 */
	while (g_pObj != NULL) {
	    DEB(("Cleaning out escaped object |%lx|\n", g_pObj));

	    if (g_pObj->pDispatch != NULL)
		g_pObj->pDispatch->Release();

	    if (g_pObj->pTypeInfo != NULL)
		g_pObj->pTypeInfo->Release();

	    g_pObj = g_pObj->pNext;
	}
	/* XXX Do we need a similar list for Win32::OLE::Variant objects? */
	OleUninitialize();
	break;

    default:
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
    SV **sp = mark + items;

    if (items == 0)
	return FALSE;

    if (!sv_isobject(ST(0)) || !sv_derived_from(ST(0), "Win32::OLE"))
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
    HV *stash; /* class for new object */
    CLSID CLSIDObj;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    unsigned int length;
    char *buffer;
    HKEY handle;
    IDispatch *pDispatch;

    if (CallObjectMethod(mark,ax,items,"new"))
	return;

    if (items < 2 || items > 3)
	DoCroak("Usage: Win32::OLE->new(class[,destroy])");

    SV *self = ST(0);
    SV *oleclass = ST(1);
    SV *destroy = NULL;

    ST(0) = &sv_undef;

    if (SvROK(self))
	stash = SvSTASH(SvRV(self));
    else
	stash = gv_stashsv(self, TRUE);

    if (items == 3) {
	destroy = ST(2);
	if (SvPOK(destroy))
	    destroy = newSVsv(destroy);
	else if (SvROK(destroy) && SvTYPE(SvRV(destroy)) == SVt_PVCV)
	    destroy = newRV_inc(SvRV(destroy));
	else { /* now it MUST be C<undef> */
	    if (SvOK(destroy))
		DoCroak("Win32::OLE::new: optional 'destroy' parameter "
			"MUST be a method name or a CODE reference");
	    destroy = NULL;
	}
    }

    buffer = SvPV(oleclass, length);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ);
    LastOleError = CLSIDFromProgID(pBuffer, &CLSIDObj);
    ReleaseBuffer(pBuffer, Buffer);

    if (!CheckOleError(LastOleError, NULL, NULL)) {
	LastOleError = CoCreateInstance(CLSIDObj, NULL, CLSCTX_LOCAL_SERVER,
					IID_IDispatch, (void**)&pDispatch);
	if (FAILED(LastOleError)) {
	    LastOleError = CoCreateInstance(CLSIDObj, NULL, CLSCTX_ALL,
					    IID_IDispatch, (void**)&pDispatch);
	    CheckOleError(LastOleError, NULL, NULL);
	}

	if (SUCCEEDED(LastOleError)) {
	    ST(0) = CreatePerlObject(stash, pDispatch, destroy);
	    DEB(("Win32::OLE::new |%lx| |%lx|\n", ST(0), pDispatch));
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

    DEB(("Win32::OLE::DESTROY |%lx| |%lx|\n", pObj, pObj->destroy));
    if (pObj->destroy != NULL) {
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
Dispatch(self,funcName,funcReturn,...)
    SV *self
    SV *funcName
    SV *funcReturn
PPCODE:
{
    char *buffer;
    char *ptr;
    unsigned int length, argErr;
    int index, arrayIndex, baseIndex;
    I32 len;
    WINOLEOBJECT *pObj;
    EXCEPINFO excepinfo;
    DISPID dispID;
    VARIANT result;
    DISPPARAMS dispParams;
    SV *curitem, *sv;
    HE **rghe = NULL; /* named argument names */

    if (!sv_isobject(self))
	DoCroak("Win32::OLE::Dispatch cannot be called as class method");

    ST(0) = &sv_no;

    pObj = GetOleObject(self);
    VariantInit(&result);
    baseIndex = 0;
    buffer = SvPV(funcName, length);
    DEB(("Dispatch \"%s\"\n", buffer));
    if (!GetHashedDispID(pObj, buffer, length, dispID)) {
	/* if the name was not found then try it as a parameter */
	/* to the default dispID */
	baseIndex = 1;
	dispID = DISPID_VALUE;
    }

    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = items - 3 + baseIndex;

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

	New(2101, rghe, dispParams.cNamedArgs, HE *);
	New(2101, rgszNames, 1+dispParams.cNamedArgs, OLECHAR *);
	New(2101, rgdispids, 1+dispParams.cNamedArgs, DISPID);
	New(2101, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	New(2101, dispParams.rgdispidNamedArgs, dispParams.cNamedArgs, DISPID);

	rgszNames[0] = AllocOleString(buffer, length);
	hv_iterinit(hv);
	for (index = 0; index < dispParams.cNamedArgs; ++index) {
	    rghe[index] = hv_iternext(hv);
	    char *pszName = hv_iterkey(rghe[index], &len);
	    rgszNames[1+index] = AllocOleString(pszName, len);
	}

	LastOleError = pObj->pDispatch->GetIDsOfNames(IID_NULL, rgszNames,
			      1+dispParams.cNamedArgs, lcidDefault, rgdispids);
	if (FAILED(LastOleError)) {
	    SV *sv = sv_2mortal(newSVpv("",0));
	    unsigned int cErrors = 0;
	    unsigned int error = 0;

	    for (index = 1 ; index <= dispParams.cNamedArgs ; ++index)
		if (rgdispids[index] == DISPID_UNKNOWN)
		   ++cErrors;

	    for (index = 1 ; index <= dispParams.cNamedArgs ; ++index)
		if (rgdispids[index] == DISPID_UNKNOWN) {
		    if (error++ > 0)
			sv_catpv(sv, error == cErrors ? " and " : ", ");
		    sv_catpvf(sv, "\"%s\"", hv_iterkey(rghe[index-1], &len));
		}

	    sv_catpvf(sv, " in methodcall/getproperty \"%s\"", buffer);
	    CheckOleError(LastOleError, NULL, sv);
	}

	for (index = 0; index <= dispParams.cNamedArgs; ++index) {
	    SysFreeString(rgszNames[index]);
	    if (index > 0 && SUCCEEDED(LastOleError)) {
		dispParams.rgdispidNamedArgs[index-1] = rgdispids[index];
		SetVariantFromSV(hv_iterval(hv, rghe[index-1]), 
				 &dispParams.rgvarg[index-1]);
	    }
	}
	Safefree(rgszNames);
	Safefree(rgdispids);

	if (FAILED(LastOleError)) {
	    Safefree(dispParams.rgvarg);
	    Safefree(dispParams.rgdispidNamedArgs);
	    XSRETURN(1);
	}

	--items;
    }

    if (dispParams.cArgs > dispParams.cNamedArgs) {
	if (dispParams.rgvarg == NULL)
	    New(2101, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);

	for(index = dispParams.cNamedArgs;
	    index < dispParams.cArgs - baseIndex;
	    ++index)
	{
	    SetVariantFromSV(ST(items-1-(index-dispParams.cNamedArgs)),
				&dispParams.rgvarg[index]);
	}

	if (baseIndex != 0)
	    SetVariantFromSV(ST(1),
				&dispParams.rgvarg[dispParams.cArgs-1]);
    }

    memset(&excepinfo, 0, sizeof(EXCEPINFO));
    LastOleError = pObj->pDispatch->Invoke(dispID, IID_NULL, lcidDefault,
				    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				    &dispParams, &result, &excepinfo, &argErr);

    if (FAILED(LastOleError)) {
	/* mega kludge. if a method in WORD is called and we ask
	 * for a result when one is not returned then
	 * hResult == DISP_E_EXCEPTION. this only happens on
	 * functions whose DISPID > 0x8000 */

	if (LastOleError == DISP_E_EXCEPTION && dispID > 0x8000) {
	    memset(&excepinfo, 0, sizeof(EXCEPINFO));
	    VariantClear(&result);
	    LastOleError = pObj->pDispatch->Invoke(dispID, IID_NULL, lcidDefault,
				    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				    &dispParams, NULL, &excepinfo, &argErr);
	}
    }

    if (FAILED(LastOleError)) {
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "in methodcall/getproperty \"%s\"", buffer);
	if (LastOleError == DISP_E_TYPEMISMATCH ||
	    LastOleError == DISP_E_PARAMNOTFOUND) /* already caught above? */
	{
	    if (argErr < dispParams.cNamedArgs)
		sv_catpvf(sv, " argument \"%s\"", hv_iterkey(rghe[argErr], &len));
	    else
		sv_catpvf(sv, " argument %d", 1 + dispParams.cArgs - argErr);
	}

	CheckOleError(LastOleError, &excepinfo, sv);
    }
    else {
	ST(0) = &sv_yes;
	SetSVFromVariant(&result, funcReturn, pObj->stash);
    }

    VariantClear(&result);
    if (dispParams.cArgs != 0) {
	for(index = 0; index < dispParams.cArgs; ++index)
	    VariantClear(&dispParams.rgvarg[index]);

	Safefree(dispParams.rgvarg);
	if (dispParams.cNamedArgs != 0) {
	    Safefree(rghe);
	    Safefree(dispParams.rgdispidNamedArgs);
	}
    }

    XSRETURN(1);
}

void
LastError(...)
PPCODE:
{
    if (CallObjectMethod(mark,ax,items,"LastError"))
	return;

    /* Direct function call accepted for compatibility */
    if (items > 1)
	DoCroak("Usage: Win32::OLE->LastError()");

    XSRETURN_IV(LastOleError);
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
    IUnknown *pUnknown;
    IDispatch *pDispatch;

    if (CallObjectMethod(mark,ax,items,"GetActiveObject"))
	return;

    if (items != 2)
	DoCroak("Usage: Win32::OLE->GetActiveObject(oleclass)");

    SV *self = ST(0);
    SV *oleclass = ST(1);

    if (!SvPOK(self))
	DoCroak("Win32::OLE->GetActiveObject must be called as a class method");

    buffer = SvPV(oleclass, length);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ);
    LastOleError = CLSIDFromProgID(pBuffer, &CLSIDObj);
    ReleaseBuffer(pBuffer, Buffer);
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    LastOleError = GetActiveObject(CLSIDObj, 0, &pUnknown);
    /* Don't call CheckOleError! Return "undef" for "Server not running" */
    if (FAILED(LastOleError))
	XSRETURN_UNDEF;

    LastOleError = pUnknown->QueryInterface(IID_IDispatch, (void**)&pDispatch);
    pUnknown->Release();
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(gv_stashsv(self, TRUE), pDispatch, NULL);
    DEB(("Win32::OLE::GetActiveObject |%lx| |%lx|\n", ST(0), pDispatch));
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

    if (CallObjectMethod(mark,ax,items,"GetObject"))
	return;

    if (items != 2)
	DoCroak("Usage: Win32::OLE->GetObject(pathname)");

    SV *self = ST(0);
    SV *pathname = ST(1);

    if (!SvPOK(self))
	DoCroak("Win32::OLE->GetObject must be called as a class method");

    LastOleError = CreateBindCtx(0, &pBindCtx);
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    buffer = SvPV(pathname, na);
    pBuffer = GetWideChar(buffer, Buffer, OLE_BUF_SIZ);
    LastOleError = MkParseDisplayName(pBindCtx, pBuffer, &ulEaten, &pMoniker);
    ReleaseBuffer(pBuffer, Buffer);
    if (FAILED(LastOleError)) {
	pBindCtx->Release();
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "after character %lu in \"%s\"", ulEaten, buffer);
	CheckOleError(LastOleError, NULL, sv);
	XSRETURN_UNDEF;
    }

    LastOleError = pMoniker->
	BindToObject(pBindCtx, NULL, IID_IDispatch, (void**)&pDispatch);
    pBindCtx->Release();
    pMoniker->Release();
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    ST(0) = CreatePerlObject(gv_stashsv(self, TRUE), pDispatch, NULL);
    XSRETURN(1);
}

void
QueryObjectType(...)
PPCODE:
{
    if (CallObjectMethod(mark,ax,items,"QueryObjectType"))
	return;

    if (items != 2)
	DoCroak("Usage: Win32::OLE->QueryObjectType(object)");

    SV *object = ST(1);

    if (!sv_isobject(object) || !sv_derived_from(object, "Win32::OLE"))
	XSRETURN_UNDEF;

    WINOLEOBJECT *pObj = GetOleObject(object);
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int count;
    BSTR bstr;
    char szName[64];
    char *pszName;

    LastOleError = pObj->pDispatch->GetTypeInfoCount(&count);
    if (CheckOleError(LastOleError, NULL, NULL) || count == 0)
	XSRETURN_UNDEF;

    LastOleError = pObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    /* Return ('TypeLib Name', 'Class Name') in array context */
    if (GIMME_V == G_ARRAY) {
	LastOleError = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	if (CheckOleError(LastOleError, NULL, NULL)) {
	    pTypeInfo->Release();
	    XSRETURN_UNDEF;
	}

	LastOleError = pTypeLib->GetDocumentation(-1, &bstr, NULL, NULL, NULL);
	pTypeLib->Release();
	if (CheckOleError(LastOleError, NULL, NULL)) {
	    pTypeInfo->Release();
	    XSRETURN_UNDEF;
	}

	pszName = GetMultiByte(bstr, szName, sizeof(szName));
	PUSHs(sv_2mortal(newSVpv(pszName, 0)));
	SysFreeString(bstr);
	ReleaseBuffer(pszName, szName);
    }

    LastOleError = pTypeInfo->
	GetDocumentation(MEMBERID_NIL, &bstr, NULL, NULL, NULL);
    pTypeInfo->Release();
    if (CheckOleError(LastOleError, NULL, NULL))
	XSRETURN_UNDEF;

    pszName = GetMultiByte(bstr, szName, sizeof(szName));
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

    /* unlink from list */
    if (pObj->pPrevious == NULL) {
	g_pObj = pObj->pNext;
	if (pObj->pNext != NULL)
	    pObj->pNext->pPrevious = NULL;
    }
    else if (pObj->pNext == NULL)
	pObj->pPrevious->pNext = NULL;
    else {
	pObj->pPrevious->pNext = pObj->pNext;
	pObj->pNext->pPrevious = pObj->pPrevious;
    }

    DEB(("Win32::OLE::Tie::DESTROY |%lx| |%lx|\n", pObj, pObj->pDispatch));
    DestroyPerlObject(pObj);
    XSRETURN_EMPTY;
}


void
FETCH(self,key)
    SV *self
    SV *key
PPCODE:
{
    SV **coo;
    char *buffer;
    unsigned int length, argErr;
    int baseIndex;
    WINOLEOBJECT *pObj;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    VARIANTARG propName;
    DISPID dispID;

    ST(0) = &sv_undef;

    coo = hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DEB(("Win32::OLE::Tie::FETCH |%s| |%d| |%lx|\n",
	 PERL_OLE_ID, PERL_OLE_IDLEN, coo));

    if (coo == NULL)
	DoCroak("Win32::OLE::Tie::FETCH: not a Win32::OLE object");

    buffer = SvPV(key, length);
    if (strEQ(buffer, PERL_OLE_ID)) {
	ST(0) = *coo;
	XSRETURN(1);
    }

    pObj = CheckOleStruct(SvIV(*coo));
    VariantInit(&result);
    VariantInit(&propName);

    baseIndex = 0;
    if (!GetHashedDispID(pObj, buffer, length, dispID)) {
	/* if the name was not found then try it as a parameter */
	/* to the default dispID */
	baseIndex = 1;
	dispID = DISPID_VALUE;
    }

    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = baseIndex;

    if (baseIndex != 0) {
	dispParams.rgvarg = &propName;
	V_VT(&propName) = VT_BSTR;
	V_BSTR(&propName) = AllocOleString(buffer, length);
    }

    memset(&excepinfo, 0, sizeof(EXCEPINFO));

    LastOleError = pObj->pDispatch->Invoke(dispID, IID_NULL,
		    lcidDefault, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
		    &dispParams, &result, &excepinfo, &argErr);

    if (FAILED(LastOleError)) {
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "in methodcall/getproperty \"%s\"", buffer);
	CheckOleError(LastOleError, &excepinfo, sv);
    }
    else
	ST(0) = SetSVFromVariant(&result, sv_newmortal(), pObj->stash);

    VariantClear(&result);
    VariantClear(&propName);

    XSRETURN(1);
}

void
STORE(self,key,value)
    SV *self
    SV *key
    SV *value
PPCODE:
{
    unsigned int length, argErr;
    char *buffer;
    int index, baseIndex;
    EXCEPINFO excepinfo;
    DISPID dispID, dispIDParam;
    DISPPARAMS dispParams;
    VARIANTARG propertyValue[2];

    WINOLEOBJECT *pObj = GetOleObject(self);

    baseIndex = 0;
    buffer = SvPV(key, length);
    if (!GetHashedDispID(pObj, buffer, length, dispID)) {
	/* if the name was not found then try it as a parameter */
	/* to the default dispID */
	baseIndex = 1;
	dispID = DISPID_VALUE;
    }

    dispIDParam = DISPID_PROPERTYPUT;
    dispParams.rgvarg = propertyValue;
    dispParams.rgdispidNamedArgs = &dispIDParam;
    dispParams.cNamedArgs = 1;
    dispParams.cArgs = 1+baseIndex;

    VariantInit(&propertyValue[0]);
    VariantInit(&propertyValue[1]);

    if (dispParams.cArgs > 0) {
	SetVariantFromSV(value, &propertyValue[0]);
	if (baseIndex != 0) {
	    V_VT(&propertyValue[1]) = VT_BSTR;
	    V_BSTR(&propertyValue[1]) = AllocOleString(buffer, length);
	}
    }

    memset(&excepinfo, 0, sizeof(EXCEPINFO));
    LastOleError = pObj->pDispatch->Invoke(dispID, IID_NULL,
				    lcidDefault, DISPATCH_PROPERTYPUT,
				    &dispParams, NULL, &excepinfo, &argErr);

    if (FAILED(LastOleError)) {
	SV *sv = sv_newmortal();
	sv_setpvf(sv, "in setproperty \"%s\"", buffer);
	CheckOleError(LastOleError, &excepinfo, sv);
    }

    for(index = 0; index < dispParams.cArgs; ++index)
	VariantClear(&propertyValue[index]);
}


void
FIRSTKEY(self)
    SV *self
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(self);
    FetchTypeInfo(pObj);
    if (pObj->pTypeInfo == NULL)
	ST(0) = &sv_undef;
    else {
	pObj->PropIndex = 0;
	ST(0) = NextPropertyName(pObj);
	if (!SvREADONLY(ST(0)))
	    sv_2mortal(ST(0));
    }
    XSRETURN(1);
}

void
NEXTKEY(self,lastKey)
    SV *self
    SV *lastKey
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(self);
    ST(0) = NextPropertyName(pObj);
    if (!SvREADONLY(ST(0)))
	sv_2mortal(ST(0));
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Const

void
_Load(clsid,major,minor,lcid,tlb)
    SV *clsid
    IV major
    IV minor
    IV lcid
    SV *tlb
PPCODE:
{
    ITypeLib *pTypeLib;
    LPTYPEATTR pTypeAttr;
    CLSID CLSIDObj;
    OLECHAR Buffer[OLE_BUF_SIZ];
    OLECHAR *pBuffer;
    char *pszClassname = "Win32::OLE";
    HV *stash = gv_stashpv(pszClassname, TRUE);

    if (sv_derived_from(clsid, pszClassname)) {
	/* Get containing typelib from IDispatch interface */
	WINOLEOBJECT *pObj = GetOleObject(clsid);
	ITypeInfo *pTypeInfo;
	unsigned int count;

	if (FAILED(pObj->pDispatch->GetTypeInfoCount(&count)))
	    DoCroak("Win32::OLE::CONST::_Load: GetTypeInfoCount failed\n");

	if (count == 0)
	    XSRETURN_UNDEF;

	LastOleError = pObj->pDispatch->GetTypeInfo(0, lcidDefault, &pTypeInfo);
	if (CheckOleError(LastOleError, NULL, NULL))
	    XSRETURN_UNDEF;

	LastOleError = pTypeInfo->GetContainingTypeLib(&pTypeLib, &count);
	pTypeInfo->Release();
	if (CheckOleError(LastOleError, NULL, NULL))
	    XSRETURN_UNDEF;
    }
    else {
	/* try to load registered typelib by clsid, version and lcid */
	char *pszBuffer = SvPV(clsid, na);
	pBuffer = GetWideChar(pszBuffer, Buffer, OLE_BUF_SIZ);
	LastOleError = CLSIDFromString(pBuffer, &CLSIDObj);
	ReleaseBuffer(pBuffer, Buffer);

	if (CheckOleError(LastOleError, NULL, NULL))
	    XSRETURN_UNDEF;

	LastOleError = LoadRegTypeLib(CLSIDObj, major, minor, lcid, &pTypeLib);
	if (FAILED(LastOleError) && SvPOK(tlb)) {
	    /* typelib not registerd, try to read from file "tlb" */
	    pszBuffer = SvPV(tlb, na);
	    pBuffer = GetWideChar(pszBuffer, Buffer, OLE_BUF_SIZ);
	    LastOleError = LoadTypeLib(pBuffer, &pTypeLib);
	    ReleaseBuffer(pBuffer, Buffer);
	}
	if (CheckOleError(LastOleError, NULL, NULL))
	    XSRETURN_UNDEF;
    }

    /* we'll return ref to hash with constant name => value pairs */
    HV *hv = newHV();
    unsigned int count = pTypeLib->GetTypeInfoCount();

    /* loop through all objects in type lib */
    for (int index=0 ; index < count ; ++index) {
	ITypeInfo *pTypeInfo;

	LastOleError = pTypeLib->GetTypeInfo(index, &pTypeInfo);
	if (CheckOleError(LastOleError, NULL, NULL))
	    continue;

	LastOleError = pTypeInfo->GetTypeAttr(&pTypeAttr);
	if (CheckOleError(LastOleError, NULL, NULL)) {
	    pTypeInfo->Release();
	    continue;
	}

	/* extract all constants for each ENUM */
	if (pTypeAttr->typekind == TKIND_ENUM) {
	    for (int iVar=0 ; iVar < pTypeAttr->cVars ; ++iVar) {
		LPVARDESC pVarDesc;

		LastOleError = pTypeInfo->GetVarDesc(iVar, &pVarDesc);
		if (CheckOleError(LastOleError, NULL, NULL))
		    continue;

		if (pVarDesc->varkind == VAR_CONST &&
		    !(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
					     VARFLAG_FRESTRICTED |
					     VARFLAG_FNONBROWSABLE))) {
		    unsigned int cName;
		    BSTR bstr;
		    char szName[64];

		    LastOleError = pTypeInfo->GetNames(pVarDesc->memid,
						       &bstr, 1, &cName);
		    if (CheckOleError(LastOleError, NULL, NULL)
			|| cName == 0 || bstr == NULL)
			continue;

		    char *pszName = GetMultiByte(bstr, szName, sizeof(szName));
		    SV *sv = newSVpv("", 0);
		    SetSVFromVariant(pVarDesc->lpvarValue, sv, stash);
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

    ST(0) = sv_2mortal(newRV_noinc((SV*)hv));
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Enum

void
_NewEnum(object)
    SV *object
PPCODE:
{
    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    IEnumVARIANT *pEnum;
    IUnknown *punk;

    WINOLEOBJECT *pObj = GetOleObject(object);

    VariantInit(&result);

    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = 0;

    memset(&excepinfo, 0, sizeof(EXCEPINFO));

    LastOleError = pObj->pDispatch->Invoke(DISPID_NEWENUM, IID_NULL,
			    lcidDefault, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			    &dispParams, &result, &excepinfo, &argErr);
    if (CheckOleError(LastOleError, &excepinfo, NULL) ||
	(V_VT(&result)&~VT_BYREF) != VT_UNKNOWN)
	DoCroak("Win32::OLE::Enum::_NewEnum: didn't return IUnknown interface");

    if (V_ISBYREF(&result))
	punk = *V_UNKNOWNREF(&result);
    else
	punk = V_UNKNOWN(&result);

    LastOleError = punk->QueryInterface(IID_IEnumVARIANT, (void**)&pEnum);
    if (CheckOleError(LastOleError, NULL, NULL))
	DoCroak("Win32::OLE::Enum::_NewEnum: "
		"missing IEnumVARIANT interface support");

    VariantClear(&result);
    XSRETURN_IV((I32)pEnum);
}

void
_Clone(pEnum)
    IEnumVARIANT *pEnum
PPCODE:
{
    IEnumVARIANT *pClone = NULL;
    LastOleError = pEnum->Clone(&pClone);
    CheckOleError(LastOleError, NULL, NULL);
    XSRETURN_IV((I32)pClone);
}

void
_Next(pEnum,object)
    IEnumVARIANT *pEnum
    SV           *object
PPCODE:
{
    WINOLEOBJECT *pObj = GetOleObject(object);
    VARIANT result;

    ST(0) = &sv_undef;
    VariantInit(&result);
    if (pEnum->Next(1, &result, NULL) == S_OK)
	ST(0) = SetSVFromVariant(&result, sv_newmortal(), pObj->stash);
    VariantClear(&result);

    XSRETURN(1);
}

void
_Release(pEnum)
    IEnumVARIANT *pEnum
PPCODE:
{
    LastOleError = pEnum->Release();
    CheckOleError(LastOleError, NULL, NULL);
    ST(0) = (LastOleError == S_OK) ? &sv_yes : &sv_no;
    XSRETURN(1);
}

void
_Reset(pEnum)
    IEnumVARIANT *pEnum
PPCODE:
{
    LastOleError = pEnum->Reset();
    CheckOleError(LastOleError, NULL, NULL);
    ST(0) = (LastOleError == S_OK) ? &sv_yes : &sv_no;
    XSRETURN(1);
}

void
_Skip(pEnum,ulCount)
    IEnumVARIANT  *pEnum
    unsigned long ulCount
PPCODE:
{
    LastOleError = pEnum->Skip(ulCount);
    CheckOleError(LastOleError, NULL, NULL);
    ST(0) = (LastOleError == S_OK) ? &sv_yes : &sv_no;
    XSRETURN(1);
}

##############################################################################

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Variant

void
new(self,type,data)
    SV *self
    SV *type
    SV *data
PPCODE:
{
    VARIANT variant;
    char *ptr;
    STRLEN length;

    VariantInit(&variant);
    V_VT(&variant) = SvIV(type);

    if (V_ISBYREF(&variant))
	DoCroak("Win32::OLE::Variant::new: VT_BYREF is not supported");

    /* XXX requirement to call mg_get() may change in Perl > 5.004 */
    if (SvGMAGICAL(data))
	mg_get(data);

    switch (V_VT(&variant)) {
    case VT_EMPTY:
    case VT_NULL:
	break;

    case VT_I2:
	V_I2(&variant) = SvIV(data);
	break;

    case VT_I4:
	V_I4(&variant) = SvIV(data);
	break;

    case VT_R4:
	V_R4(&variant) = SvNV(data);
	break;

    case VT_R8:
	V_R8(&variant) = SvNV(data);
	break;

    case VT_CY:
    case VT_DATE:
	V_VT(&variant) = VT_BSTR;
	ptr = SvPV(data, length);
	V_BSTR(&variant) = AllocOleString(ptr, length);
	VariantChangeType(&variant, &variant, 0, SvIV(type));
	break;

    case VT_BSTR:
	ptr = SvPV(data, length);
	V_BSTR(&variant) = AllocOleString(ptr, length);
	break;

    case VT_DISPATCH:
	/* Argument MUST be a valid Perl OLE object! */
	V_DISPATCH(&variant) = GetOleObject(data)->pDispatch;
	V_DISPATCH(&variant)->AddRef();
	break;

    case VT_ERROR:
	V_ERROR(&variant) = SvIV(data);
	break;

    case VT_BOOL:
	/* Either all bits are 0 or ALL bits MUST BE 1 */
	V_BOOL(&variant) = SvIV(data) ? ~0 : 0;
	break;

    /* case VT_VARIANT: invalid without VT_BYREF */

    case VT_UNKNOWN:
	/* Argument MUST be a valid Perl OLE object! */
	/* Query IUnknown interface to allow identity tests */
	LastOleError = GetOleObject(data)->pDispatch->
	    QueryInterface(IID_IUnknown, (void**)&V_UNKNOWN(&variant));
	CheckOleError(LastOleError, NULL, NULL);
	break;

    case VT_UI1:
	if (SvPOK(data)) {
	    unsigned char* pDest;

	    ptr = SvPV(data, length);
	    V_ARRAY(&variant) = SafeArrayCreateVector(VT_UI1, 0, length);
	    if (V_ARRAY(&variant) != NULL) {
		V_VT(&variant) = VT_UI1 | VT_ARRAY;
		LastOleError = SafeArrayAccessData(V_ARRAY(&variant),
						   (void**)&pDest);
		if (!CheckOleError(LastOleError, NULL, NULL)) {
		    memcpy(pDest, ptr, length);
		    SafeArrayUnaccessData(V_ARRAY(&variant));
		}
	    }
	}
	else
	    V_UI1(&variant) = SvIV(data);

	break;

    default:
	DoCroak("Win32::OLE::Variant::new: invalid value type %d", 
		V_VT(&variant));
    }

    SV *sv = newSVpv((char*)&variant, sizeof(variant));
    HV *stash;

    if (SvROK(self))
	stash = SvSTASH(SvRV(self));
    else
	stash = gv_stashsv(self, TRUE);

    ST(0) = sv_2mortal(sv_bless(newRV_noinc(sv), stash));
    XSRETURN(1);
}

void
DESTROY(self)
    SV *self
PPCODE:
{
    STRLEN len;
    VARIANT *pVariant = (VARIANT*)SvPV(SvRV(self), len);

    if (len != sizeof(VARIANT))
	DoCroak("Win32::OLE::Variant::DESTROY: Invalid object");

    VariantClear(pVariant);
    XSRETURN_EMPTY;
}
