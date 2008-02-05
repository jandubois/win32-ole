/* OLE.xs
 *
 * (c) 1995 Microsoft Corporation. All rights reserved. 
 *	Developed by ActiveWare Internet Corp., http://www.ActiveWare.com
 *
 *    You may distribute under the terms of either the GNU General Public
 *    License or the Artistic License, as specified in the README file.
 */

#include <stdlib.h>
#include <math.h>	/* this hack gets around VC-5.0 brainmelt */
#include <windows.h>
#include <objbase.h>
#include <winnt.h>
#include <oleauto.h>
#include <malloc.h>
#define WIN32_LEAN_AND_MEAN

#if defined(__cplusplus)
extern "C" {
#endif
#include "EXTERN.h"
#include "perl.h"
#include "XSub.h"

#undef bool   /* perl defines bool to char which breaks things */

// #define MYDEBUG

#undef DEB

#ifdef MYDEBUG
#define DEB(a)		a
#else
#define DEB(a)
#endif

#define SUCCESSRETURNED(x)	(x == ERROR_SUCCESS)
#define RETURNRESULT if (bSuccess) { XST_mYES(0); }\
		     else	   { XST_mNO(0); }\
		     XSRETURN(1)

#define NEW(x,v,n,t)  (v = (t*)safemalloc((MEM_SIZE)((n) * sizeof(t))))
#define INTERNAL_ERROR ((HRESULT)-1)

static long LastOLEError;
static const DWORD WINOLE_MAGIC = 0x12344321;
static const int OLE_BUF_SIZ = 1024;
static const LCID lcidDefault = (0x02 << 10) /* LOCALE_SYSTEM_DEFAULT */;
static char PERL_OLE_ID[] = "___Perl___OLEObject___";
static const int PERL_OLE_IDLEN = sizeof(PERL_OLE_ID)-1;
static char IV_TYPE[] = "Type";
static const int IV_TYPELEN = sizeof(IV_TYPE)-1;
static char IV_VAL[] = "Value";
static const int IV_VALLEN = sizeof(IV_VAL)-1;

typedef struct _tagWINOLEOBJECT *LPWINOLEOBJECT;
typedef struct _tagWINOLEOBJECT
{
    long Win32OLEMagic;
    LPWINOLEOBJECT pNext;
    LPWINOLEOBJECT pPrevious;
    IDispatch*	pDispatch;
    IEnumVARIANT *pEnum;
    HV*	hashTable;
} WINOLEOBJECT; 

static LPWINOLEOBJECT g_lpObj = NULL;

#ifdef _DEBUG

inline void
EXCEPTIONINFOCLEAR(EXCEPINFO &excepinfo)
{
    memset(&excepinfo, 0, sizeof(EXCEPINFO));
}

inline void
EXCEPTIONINFO(HRESULT lastError, EXCEPINFO &excepinfo)
{
    if (FAILED(lastError)) {
	OutputDebugString("Exception Source: ");
	ODSOLE(excepinfo.bstrSource);
	OutputDebugString("	Description: ");
	ODSOLE(excepinfo.bstrDescription);
    }
}

inline void
ODS(LPSTR x)
{
    OutputDebugString(x);
    OutputDebugString("\n");
}

#if defined(UNICODE)
#    define ODSOLE(x) ODS(x)
#else
inline void
ODSOLE(LPCWSTR x)
{
    char bufA[256];

    if (x != NULL)
	WideCharToMultiByte(CP_ACP, NULL, x, -1, bufA, 256, NULL, NULL);
    else
	strcpy(bufA, "<null>");

    ODS(bufA);
}
#endif	/* UNICODE */

#else	/* _DEBUG */

#    define EXCEPTIONINFOCLEAR(x)
#    define EXCEPTIONINFO(x, y)
#    define ODS(x)
#    define ODSOLE(x)

#endif	/* _DEBUG */


void
ReleaseObjects(LPWINOLEOBJECT lpObj)
{
    DEB(fprintf(stderr, "ReleaseObjects |%lx|", lpObj));
    if (lpObj->pDispatch != NULL) {
	lpObj->pDispatch->Release();
	lpObj->pDispatch = NULL;
	DEB(fprintf(stderr, " pDispatch"));
    }

    if (lpObj->hashTable != NULL) {
	DEB(fprintf(stderr, " hashTable(%d)", SvREFCNT(lpObj->hashTable)));
	SvREFCNT_dec(lpObj->hashTable);
	lpObj->hashTable = NULL;
    }
    DEB(fprintf(stderr, "\n"));
}

LPWINOLEOBJECT
NewDispatch(IDispatch* pDisp, BOOL bCreated)
{
    LPWINOLEOBJECT lpObj;
    NEW(2101, lpObj, 1, WINOLEOBJECT);
    lpObj->Win32OLEMagic = WINOLE_MAGIC;
    lpObj->pPrevious = NULL;
    lpObj->pEnum = NULL;
    lpObj->pDispatch = pDisp;
    lpObj->hashTable = newHV();
    lpObj->pNext = g_lpObj;
    if (g_lpObj)
	g_lpObj->pPrevious = lpObj;
    g_lpObj = lpObj;
    DEB(fprintf(stderr, "NewDispatch = |%lx|\n", lpObj));

    if (!bCreated)
	pDisp->AddRef();

    return lpObj;
}

/* Converts dest into an RV pointing to obj. Doesn't increment
 * refcount of obj.  The RV will be blessed if classname is non-null */
SV *
sv_setrv(SV *dest, SV *obj, char *classname)
{
    sv_upgrade(dest, SVt_RV);
    SvRV(dest) = obj;
    SvTEMP_off(obj);
    SvROK_on(dest);
    if (classname)
	(void)sv_bless(dest, gv_stashpv(classname, TRUE));
    return dest;
}

/* converts newref into an RV that points to a new perl OLE object */
IV
CreatePerlObject(SV *newref, IDispatch *pDisp, BOOL bCreated)
{
    if (pDisp != NULL) {
	HV *hvouter = newHV();
	HV *hvinner = newHV();
	SV *inner = newSVpv("",0);

	hv_store(hvinner, PERL_OLE_ID, PERL_OLE_IDLEN,
		 newSViv((long)NewDispatch(pDisp, bCreated)), 0);

	sv_setrv(inner, (SV*)hvinner, "Win32::OLE::Tie");
	sv_magic((SV*)hvouter, inner, 'P', Nullch, 0);
	SvREFCNT_dec(inner);

	sv_setrv(newref, (SV*)hvouter, "Win32::OLE");
	return TRUE;
    }
    return FALSE;
}

LPWINOLEOBJECT
GetOLEObject(SV *sv) 
{
    SV **psv;

    if (sv != NULL && SvROK(sv)) {
	psv = hv_fetch((HV*)SvRV(sv), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
	if (psv != NULL) {
	    DEB(fprintf(stderr, "GetOLEObject = |%lx|\n", SvIV(*psv)));
	    return (LPWINOLEOBJECT)SvIV(*psv);
	}
    }
    return (LPWINOLEOBJECT)NULL;
}

inline BOOL
IsOleStruct(LPWINOLEOBJECT lpObj)
{
    return (lpObj != NULL && lpObj->Win32OLEMagic == WINOLE_MAGIC);
}

inline BOOL
ValidDispatch(LPWINOLEOBJECT lpObj)
{
    return (IsOleStruct(lpObj) && lpObj->pDispatch != NULL);
}

BSTR
AllocOLEString(char* lpStr, int length)
{
    int count = (length+1)*2;
    OLECHAR* pOLEChar = (OLECHAR*)alloca(count);
    MultiByteToWideChar(CP_ACP, 0, lpStr, -1, pOLEChar, count);
    return SysAllocString(pOLEChar);
}



BOOL
GetHashedDispID(LPWINOLEOBJECT lpObj, char *buffer, unsigned int length, DISPID &dispID)
{
    if (length == 0 || *buffer == '\0') {
	dispID = DISPID_VALUE;
	return TRUE;
    }

    if (!hv_exists(lpObj->hashTable, buffer, length)) {
	/* not there so get if info and add it */
	DISPID id;
	OLECHAR bBuffer[OLE_BUF_SIZ], *pbBuffer;
	SV* sv;
	pbBuffer = bBuffer;
	MultiByteToWideChar(CP_ACP, NULL, buffer, -1, bBuffer, sizeof(bBuffer));
	LastOLEError = lpObj->pDispatch->GetIDsOfNames(IID_NULL, &pbBuffer, 1, lcidDefault, &id);
	if (SUCCEEDED(LastOLEError)) {
	    sv = newSViv(id);
	    hv_store(lpObj->hashTable, buffer, length, sv, 0);
			dispID = id;
			return TRUE;
	}
    }
    else {
	SV** ppsv;
	ppsv = hv_fetch(lpObj->hashTable, buffer, length, 0);
	if (ppsv != NULL) {
	    dispID = (DISPID)SvIV(*ppsv);
	    return TRUE;
	}
	LastOLEError = INTERNAL_ERROR;
    }
    return FALSE;
}

void
CreateSafeBinaryArray(SV* sv, VARIANT *pVariant)
{
    unsigned char* ptr;
    unsigned int length;
    SAFEARRAYBOUND rgsabound;
    unsigned char* pDest;

    ptr = (unsigned char*)SvPV(sv, length);
    rgsabound.lLbound = 0;
    rgsabound.cElements = length;
    pVariant->parray = SafeArrayCreate(VT_UI1, 1, &rgsabound);
    if (pVariant->parray != NULL) {
	pVariant->vt = VT_UI1 | VT_ARRAY;
	HRESULT hResult = SafeArrayAccessData(pVariant->parray, (void**)&pDest);
	if (SUCCEEDED(hResult)) {
	    memcpy(pDest, ptr, length);
	    SafeArrayUnaccessData(pVariant->parray);
	}
    }
}

void
CreateSafeArray(AV* av, VARIANT *pVariant)
{
    SV **psv;
    char *ptr;
    unsigned int length;
    SAFEARRAYBOUND rgsabound;
    long arrayIndex;
    VARIANT variant;

    rgsabound.lLbound = 0;
    rgsabound.cElements = av_len(av)+1;
    pVariant->parray = SafeArrayCreate(VT_VARIANT, 1, &rgsabound);
    if (pVariant->parray != NULL) {
	pVariant->vt = VT_VARIANT | VT_ARRAY;
	VariantInit(&variant);
	variant.vt = VT_BSTR;
	for(arrayIndex = 0; arrayIndex < rgsabound.cElements; ++arrayIndex) {
	    psv = av_fetch(av, arrayIndex, 0);
	    if (psv != NULL) {
		ptr = SvPV(*psv, length);
		variant.bstrVal = AllocOLEString(ptr, length);
		if (variant.bstrVal != NULL)
		    SafeArrayPutElement(pVariant->parray, &arrayIndex, &variant);
	    }
	}
    }
}

void
DestroySafeArray(VARIANT *pVariant)
{
    long arrayIndex, upperBound;
    VARIANT variant;
    HRESULT hResult;

    hResult = SafeArrayGetUBound(pVariant->parray, 1, &upperBound);
    if (SUCCEEDED(hResult)) {
	for(arrayIndex = 0; arrayIndex <= upperBound; ++arrayIndex) {
	    hResult = SafeArrayGetElement(pVariant->parray, &arrayIndex, &variant);
	    if (SUCCEEDED(hResult)) {
		if (variant.vt == VT_BSTR) {
		    SysFreeString(variant.bstrVal);
		    variant.bstrVal = NULL;
		    SafeArrayPutElement(pVariant->parray, &arrayIndex, &variant);
		}
	    }
	}
    }
    SafeArrayDestroy(pVariant->parray);
    pVariant->parray = NULL;
}


void
CreateVariantFromInternalVariant(SV* sv, VARIANT *pVariant)
{
    char *ptr;
    unsigned int length;
    int nType;
    SV **psv;
    HV* hv = (HV*)SvRV(sv);

    psv = hv_fetch(hv, IV_TYPE, IV_TYPELEN, 0);
    if (psv != NULL) {
	nType = SvIV(*psv);
	psv = hv_fetch(hv, IV_VAL, IV_VALLEN, 0);
	if (psv != NULL) {
	    switch(nType) {
		case VT_UI1:
		    switch(SvTYPE(*psv)) {
			case SVt_PVIV:
			case SVt_PV:
			    CreateSafeBinaryArray(*psv, pVariant);
			    break;

			default:
			    V_VT(pVariant) = VT_UI1;
			    V_UI1(pVariant) = SvIV(*psv);
			    break;
		    }
		    break;

		case VT_BOOL:
		    V_VT(pVariant) = VT_BOOL;
		    V_BOOL(pVariant) = SvIV(*psv);
		    break;

		case VT_I2:
		    V_VT(pVariant) = VT_I2;
		    V_I2(pVariant) = SvIV(*psv);
		    break;

		case VT_I4:
		    V_VT(pVariant) = VT_I4;
		    V_I4(pVariant) = SvIV(*psv);
		    break;

		case VT_R4:
		    V_VT(pVariant) = VT_R4;
		    V_R4(pVariant) = SvNV(*psv);
		    break;

		case VT_R8:
		    V_VT(pVariant) = VT_R8;
		    V_R8(pVariant) = SvNV(*psv);
		    break;

		case VT_BSTR:
		    V_VT(pVariant) = VT_BSTR;
		    ptr = SvPV(*psv, length);
		    V_BSTR(pVariant) = AllocOLEString(ptr, length);
		    break;

		case VT_DATE:
		case VT_CY:
		    V_VT(pVariant) = VT_BSTR;
		    ptr = SvPV(*psv, length);
		    V_BSTR(pVariant) = AllocOLEString(ptr, length);
		    VariantChangeType(pVariant, pVariant, 0, nType);
		    break;
	    }
	}
    }
}

void
CreateVariantFromSV(SV* sv, VARIANT *pVariant)
{
    char *ptr;
    unsigned int length;
    int type;

    VariantInit(pVariant);

    if (SvROK(sv)) {
	if (sv_isa(sv, "Win32::OLE")) {
	    LPWINOLEOBJECT lpObj;
	    lpObj = GetOLEObject(sv);
	    if (ValidDispatch(lpObj) && lpObj->hashTable != NULL) {
		V_VT(pVariant) = VT_DISPATCH;
		V_DISPATCH(pVariant) = lpObj->pDispatch;
		return;
	    }
	}
	else if (sv_isa(sv, "Win32::OLE::Variant")) {
	    CreateVariantFromInternalVariant(sv, pVariant);
	    return;
	}
	sv = SvRV(sv);
    }

    type = SvTYPE(sv);
    if (type == SVt_PVMG) {	/* blessed scalar */
	if (SvPOKp(sv))
	    type = SVt_PV;
	else if (SvNOKp(sv))
	    type = SVt_NV;
	else if (SvIOKp(sv))
	    type = SVt_IV;
    }

    switch(type)
    {
	case SVt_PVAV:
	    CreateSafeArray((AV*)sv, pVariant);
	    break;

	case SVt_PVIV:
	case SVt_PV:
	    pVariant->vt = VT_BSTR; 
	    ptr = SvPV(sv, length);
	    pVariant->bstrVal = AllocOLEString(ptr, length);
	    break;

	case SVt_NV:
	    pVariant->vt = VT_R8;
	    pVariant->dblVal = SvNV(sv);
	    break;

	default:
	    pVariant->vt = VT_I4;
	    pVariant->lVal = SvIV(sv);
	    break;
    }
}

#define SETiVRETURN(x,f)\
		    if (x->vt&VT_BYREF) {\
			sv_setiv(sv, (long)*x->p##f);\
		    } else {\
			sv_setiv(sv, (long)x->f);\
		    }

#define SETnVRETURN(x,f)\
		    if (x->vt&VT_BYREF) {\
			sv_setnv(sv, (double)*x->p##f);\
		    } else {\
			sv_setnv(sv, (double)x->f);\
		    }

void
SVFromVariant(VARIANT *pVariant, SV* sv)
{
    switch(pVariant->vt&~VT_BYREF)
    {
	case VT_EMPTY:
	case VT_NULL:
	    sv_setsv(sv, &sv_undef);
	    break;

	case VT_UI1:
	    SETiVRETURN(pVariant,bVal)
	    break;

	case VT_I2:
	    SETiVRETURN(pVariant,iVal)
	    break;

	case VT_I4:
	    SETiVRETURN(pVariant,lVal)
	    break;

	case VT_R4:
	    SETnVRETURN(pVariant,fltVal)
	    break;

	case VT_R8:
	    SETnVRETURN(pVariant,dblVal)
	    break;

	case VT_BSTR:
ConvertString:
	    {
		int length;
		char *pStr;
		if (pVariant->vt&VT_BYREF)
		    length = SysStringLen(*pVariant->pbstrVal)+2;
		else
		    length = SysStringLen(pVariant->bstrVal)+2;

		NEW(1110, pStr, length, char);

		if (pVariant->vt&VT_BYREF)
		    WideCharToMultiByte(CP_ACP, NULL, *pVariant->pbstrVal,
					-1, pStr, length, NULL, NULL);
		else
		    WideCharToMultiByte(CP_ACP, NULL, pVariant->bstrVal,
					-1, pStr, length, NULL, NULL);

		sv_setpv(sv, pStr);
		/* mega kludge -- check this in 5.003_11 */
		/* if (SvTYPE(sv) == SVt_PVNV)
		    (sv)->sv_flags &= ~SVt_NV; */

		Safefree(pStr);
	    }
	    break;

	case VT_ERROR:
	    SETiVRETURN(pVariant,scode)
	    break;

	case VT_BOOL:
	    if (pVariant->vt&VT_BYREF)
		sv_setiv(sv, (long)V_BOOLREF(pVariant));
	    else
		sv_setiv(sv, (long)V_BOOL(pVariant));
	    break;

	case VT_DISPATCH:
	    {
		IDispatch *pDisp;
		if (pVariant->vt&VT_BYREF) 
		    pDisp = (IDispatch*)*pVariant->ppunkVal;
		else
		    pDisp = (IDispatch*)pVariant->punkVal;
		CreatePerlObject(sv, pDisp,FALSE);
	    }
	    break;

	case VT_UNKNOWN:
	    {
		IUnknown *punk;
		IDispatch *pDisp;
		if (pVariant->vt&VT_BYREF) 
		    punk = (IUnknown*)*pVariant->ppunkVal;
		else
		    punk = (IUnknown*)pVariant->punkVal;
		if (SUCCEEDED(punk->QueryInterface(IID_IDispatch,
						   (void**)&pDisp))) {
		    CreatePerlObject(sv, pDisp,FALSE);
		}
		else
		    sv_setsv(sv, &sv_undef);
		punk->Release();
	    }
	    break;

	case VT_DATE:
	case VT_CY:
	case VT_VARIANT:
	default:
	    {
		HRESULT hResult = VariantChangeType(pVariant, pVariant,
						     0, VT_BSTR);
		if (SUCCEEDED(hResult))
		    goto ConvertString;
	    }

	    sv_setsv(sv, &sv_undef);
	    break;
    }
}

void
SetSVFromVariant(VARIANT *pVariant, SV* sv)
{
    if (pVariant->vt&VT_ARRAY) {	/* array of items */
	SV *nsv;
	AV *av;
	VARIANT variant;
	int dim, index;
	long *pArrayIndex, *pLowerBound, *pUpperBound;
	HRESULT hResult;

	dim = SafeArrayGetDim(pVariant->parray);
	NEW(4444, pArrayIndex, dim, long);
	NEW(4444, pLowerBound, dim, long);
	NEW(4444, pUpperBound, dim, long);
	for(index = 1; index <= dim; ++index) {
	    hResult = SafeArrayGetLBound(pVariant->parray, index,
					  &pLowerBound[index-1]);
	    if (FAILED(hResult))
		goto ErrorExit;
	}

	for(index = 1; index <= dim; ++index) {
	    hResult = SafeArrayGetUBound(pVariant->parray, index,
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
		hResult = SafeArrayGetElement(pVariant->parray, pArrayIndex,
						&variant);
		if (SUCCEEDED(hResult)) {
		    nsv = newSVpv("",0);
		    SetSVFromVariant(&variant, nsv);
		    av_push(av, nsv);
		}
	    }
	}
	sv_setrv(sv, (SV*)av, Nullch);

ErrorExit:
	Safefree(pArrayIndex);
	Safefree(pLowerBound);
	Safefree(pUpperBound);
    }
    else
	SVFromVariant(pVariant, sv);
}

SV *
Win32OLEPropertyGet(SV *object, SV *propname)
{
    char *buffer;
    unsigned int length, argErr;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    DISPID dispID;
    BOOL bSuccess = FALSE;
    SV* sv = NULL;

    lpObj = (LPWINOLEOBJECT)SvIV(object);
    if (ValidDispatch(lpObj)) {
	VariantInit(&result);

	dispParams.rgvarg = NULL;
	dispParams.rgdispidNamedArgs = NULL;
	dispParams.cNamedArgs = 0;
	dispParams.cArgs = 0;

	buffer = SvPV(propname, length);
	if (GetHashedDispID(lpObj, buffer, length, dispID)) {
	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL,
			lcidDefault, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&dispParams, &result, &excepinfo, &argErr);

	    EXCEPTIONINFO(LastOLEError, excepinfo);
	    bSuccess = SUCCEEDED(LastOLEError);
	}

	if (bSuccess) {	/* handle result */
	    sv = newSVpv("",0);
	    SetSVFromVariant(&result, sv);
	    VariantClear(&result);
	}
    }
    return sv;
}

SV *
Win32OLEPropertySet(SV *object, SV *propname, SV *val)
{
    unsigned int length, argErr;
    char *buffer;
    int index;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPID dispID, dispIDParam;
    DISPPARAMS dispParams;
    VARIANT propertyValue;
    BOOL bSuccess = FALSE;

    lpObj = (LPWINOLEOBJECT)SvIV(object);
    if (ValidDispatch(lpObj)) {
	dispIDParam = DISPID_PROPERTYPUT;
	dispParams.rgvarg = &propertyValue;
	dispParams.rgdispidNamedArgs = &dispIDParam;
	dispParams.cNamedArgs = 1;
	dispParams.cArgs = 1;

	CreateVariantFromSV(val, &propertyValue);

	buffer = SvPV(propname, length);
	if (GetHashedDispID(lpObj, buffer, length, dispID)) {
	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL,
					lcidDefault, DISPATCH_PROPERTYPUT,
					&dispParams, NULL, &excepinfo, &argErr);

	    EXCEPTIONINFO(LastOLEError, excepinfo);
	    bSuccess = SUCCEEDED(LastOLEError);
	}


	if (propertyValue.vt == VT_BSTR)
	    SysFreeString(propertyValue.bstrVal);
	else if (propertyValue.vt == (VT_VARIANT | VT_ARRAY))
	    DestroySafeArray(&propertyValue);
    }

    return (bSuccess ? val : NULL);
}

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
	/* global destruction will have normally DESTROYed all
	 * objects, so the loop below will never be entered.
	 * Unless global destruction phase was somehow interrupted.
	 */
	while (g_lpObj != NULL) {
	    LPWINOLEOBJECT i = g_lpObj;
	    DEB(fprintf(stderr, "Cleaning out escaped object |%lx|\n", i));
	    g_lpObj = i->pNext;
	    ReleaseObjects(i);
	}
	OleUninitialize();
	break;
    default:
	break;
    }
    return TRUE;
}

#if defined(__cplusplus)
}
#endif

MODULE = Win32::OLE		PACKAGE = Win32::OLE

PROTOTYPES: DISABLE


IV
LastError()
CODE:
    RETVAL = LastOLEError;
OUTPUT:
    RETVAL

IV
CreateObject(clas,obj)
    SV *clas
    SV *obj
PPCODE:
{
    CLSID CLSIDObj;
    OLECHAR bBuffer[OLE_BUF_SIZ];
    unsigned int length;
    char *buffer;
    HKEY handle;
    IDispatch *pDisp;
    BOOL bSuccess = FALSE;

    buffer = SvPV(clas, length);
    MultiByteToWideChar(CP_ACP, NULL, buffer, -1, bBuffer, sizeof(bBuffer));

    LastOLEError = CLSIDFromProgID(bBuffer, &CLSIDObj);
    if (SUCCEEDED(LastOLEError)) {
	LastOLEError = CoCreateInstance(CLSIDObj, NULL, CLSCTX_LOCAL_SERVER,
					IID_IDispatch, (void**)&pDisp);
	if (FAILED(LastOLEError))
	    LastOLEError = CoCreateInstance(CLSIDObj, NULL, CLSCTX_ALL,
					IID_IDispatch, (void**)&pDisp);
	if (SUCCEEDED(LastOLEError)) {
	    if (CreatePerlObject(obj, pDisp, TRUE)) {
		DEB(fprintf(stderr, "CreateObject = |%lx| |%lx|\n",
				(long) obj, (long)pDisp));
		bSuccess = TRUE;
	    }
	}
    }

    RETURNRESULT;
}

IV
Dispatch(object,funcName,funcReturn,...)
    SV *object
    SV *funcName
    SV *funcReturn
PPCODE:
{
    char *buffer;
    char *ptr;
    unsigned int length, argErr;
    int index, arrayIndex;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPID dispID;
    VARIANT result;
    DISPPARAMS dispParams;
    BOOL bSuccess = FALSE;

    lpObj = GetOLEObject(object);
    if (ValidDispatch(lpObj) && lpObj->hashTable != NULL) {
	VariantInit(&result);

	dispParams.rgvarg = NULL;
	dispParams.rgdispidNamedArgs = NULL;
	dispParams.cNamedArgs = 0;
	dispParams.cArgs = items - 3;

	if (dispParams.cArgs > 0) {
	    NEW(2101, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	    for(index = 0; index < dispParams.cArgs; ++index)
		CreateVariantFromSV(ST(items - 1 - index),
				    &dispParams.rgvarg[index]);
	}

	buffer = SvPV(funcName, length);
	DEB(fprintf(stderr, "Dispatch \"%s\"\n", buffer));
	if (GetHashedDispID(lpObj, buffer, length, dispID)) {
	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL,
				lcidDefault,
				DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				&dispParams, &result, &excepinfo, &argErr);

	    if (FAILED(LastOLEError)) {
	        /* mega kludge. if a method in WORD is called and we ask
		 * for a result when one is not returned then
		 * hResult == DISP_E_EXCEPTION. this only happens on
		 * functions whose DISPID > 0x8000 */
		EXCEPTIONINFO(LastOLEError, excepinfo);

		if (LastOLEError == DISP_E_EXCEPTION && dispID > 0x8000) {
		    EXCEPTIONINFOCLEAR(excepinfo);

		    VariantClear(&result);
		    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL,
					lcidDefault,
					DISPATCH_METHOD | DISPATCH_PROPERTYGET,
					&dispParams, NULL, &excepinfo, &argErr);

		    EXCEPTIONINFO(LastOLEError, excepinfo);
		    if (SUCCEEDED(LastOLEError))
			bSuccess = TRUE;
		}
	    }
	    else
		bSuccess = TRUE;
	}


	if (bSuccess) {	/* handle result */
	    SetSVFromVariant(&result, funcReturn);
	}

	VariantClear(&result);
	if (dispParams.cArgs != 0) {
	    for(index = 0; index < dispParams.cArgs; ++index) {
		if (dispParams.rgvarg[index].vt == VT_BSTR)
		    SysFreeString(dispParams.rgvarg[index].bstrVal);

		else if (dispParams.rgvarg[index].vt == (VT_VARIANT | VT_ARRAY))
		    DestroySafeArray(&dispParams.rgvarg[index]);
	    }
	    Safefree(dispParams.rgvarg);
	}
    }

    RETURNRESULT;
}

IV
GetProperty(object,varName,varReturn,...)
    SV *object
    SV *varName
    SV *varReturn
PPCODE:
{
    char *buffer;
    unsigned int length, argErr;
    int index;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    DISPID dispID;
    BOOL bSuccess = FALSE;

    lpObj = GetOLEObject(object);
    if (ValidDispatch(lpObj)) {
	VariantInit(&result);

	dispParams.rgvarg = NULL;
	dispParams.rgdispidNamedArgs = NULL;
	dispParams.cNamedArgs = 0;
	dispParams.cArgs = items - 3;
	if (dispParams.cArgs > 0) {
	    NEW(2101, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	    for(index = 0; index < dispParams.cArgs; ++index) {
		VariantInit(&dispParams.rgvarg[index]);
		dispParams.rgvarg[index].vt = VT_BSTR; 
		buffer = SvPV(ST(items - 1 - index), length);
		dispParams.rgvarg[index].bstrVal = AllocOLEString(buffer, length);
	    }
				
	}

	buffer = SvPV(varName, length);
	if (GetHashedDispID(lpObj, buffer, length, dispID)) {
	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL, lcidDefault, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
												&dispParams, &result, &excepinfo, &argErr);

	    EXCEPTIONINFO(LastOLEError, excepinfo);
	    bSuccess = SUCCEEDED(LastOLEError);
	}

	if (bSuccess) {	/* handle result */
	    SetSVFromVariant(&result, varReturn);
	    VariantClear(&result);
	}
	if (dispParams.cArgs != 0) {
	    for(index = 0; index < dispParams.cArgs; ++index) {
		if (dispParams.rgvarg[index].vt == VT_BSTR)
		    SysFreeString(dispParams.rgvarg[index].bstrVal);
	    }
	    Safefree(dispParams.rgvarg);
	}
    }

    RETURNRESULT;
}

IV
SetProperty(object,varName,varValue,...)
    SV *object
    SV *varName
    SV *varValue
PPCODE:
{
    unsigned int length, argErr;
    char *buffer;
    int index;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPID dispID;
    DISPID dispidNamed;
    DISPPARAMS dispParams;
    BOOL bSuccess = FALSE;

    lpObj = GetOLEObject(object);
    if (ValidDispatch(lpObj)) {
	dispParams.rgvarg = NULL;
	dispParams.rgdispidNamedArgs = NULL;
	dispParams.cNamedArgs = 0;
	dispParams.cArgs = items - 2;

	if (dispParams.cArgs > 0) {
	    NEW(2101, dispParams.rgvarg, dispParams.cArgs, VARIANTARG);
	    dispParams.rgvarg[0].vt = VT_BSTR; 
	    buffer = SvPV(ST(2), length);
	    dispParams.rgvarg[0].bstrVal = AllocOLEString(buffer, length);

	    if (dispParams.cArgs > 1) {
		dispidNamed = DISPID_PROPERTYPUT;
		dispParams.rgdispidNamedArgs = &dispidNamed;
		dispParams.cNamedArgs = 1;

		for(index = 1; index < dispParams.cArgs; ++index)
		    CreateVariantFromSV(ST(items - index),
					&dispParams.rgvarg[index]);
	    }
	}


	buffer = SvPV(varName, length);
	if (GetHashedDispID(lpObj, buffer, length, dispID)) {
	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(dispID, IID_NULL,
				lcidDefault, DISPATCH_PROPERTYPUT,
				&dispParams, NULL, &excepinfo, &argErr);

	    EXCEPTIONINFO(LastOLEError, excepinfo);
	    bSuccess = SUCCEEDED(LastOLEError);
	}


	if (dispParams.cArgs > 0) {
	    if (dispParams.rgvarg[0].vt == VT_BSTR)
		SysFreeString(dispParams.rgvarg[0].bstrVal);

	    Safefree(dispParams.rgvarg);
	}
    }

    RETURNRESULT;
}

MODULE = Win32::OLE		PACKAGE = Win32::OLE::Tie

void
DESTROY(self)
    SV *self
CODE:
{
    if (self && SvROK(self)) {
	LPWINOLEOBJECT lpObj = GetOLEObject(self);
	if (IsOleStruct(lpObj)) {
	    DEB(fprintf(stderr, "Win32::OLE::Tie::DESTROY |%lx| |%lx|\n",
			(long)lpObj, (long)lpObj->pDispatch));
	    ReleaseObjects(lpObj);

	    /* unlink from list */
	    if (lpObj->pPrevious == NULL) {
		g_lpObj = lpObj->pNext;
		if (lpObj->pNext != NULL)
		    lpObj->pNext->pPrevious = NULL;
	    }
	    else if (lpObj->pNext == NULL)
	        lpObj->pPrevious->pNext = NULL;
	    else {
	        lpObj->pPrevious->pNext = lpObj->pNext;
	        lpObj->pNext->pPrevious = lpObj->pPrevious;
	    }

	    Safefree(lpObj);
	}
    }
}


SV *
FETCH(self,key)
    SV *self
    SV *key
PPCODE:
{
    SV **coo;
    SV *temp;

    ST(0) = &sv_undef;
    coo = hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DEB(fprintf(stderr, "Win32::OLE::Tie::FETCH |%s| |%d| |%lx|\n",
		PERL_OLE_ID, PERL_OLE_IDLEN,(long)coo));
    if (coo != NULL) {
	if (strcmp(SvPV(key,na), PERL_OLE_ID) == 0) {
	    ST(0) = *coo;
	}
	else {
	    temp = Win32OLEPropertyGet(*coo, key);
	    if (temp && temp != &sv_undef)
		ST(0) = sv_2mortal(temp);
	}
    }
    XSRETURN(1);
}

void
STORE(self,key,value)
    SV *self
    SV *key
    SV *value
CODE:
{
    SV **coo;
    coo = hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DEB(fprintf(stderr, "Win32::OLE::Tie::STORE |%s| |%d| |%lx|\n",
		PERL_OLE_ID, PERL_OLE_IDLEN,(long)coo));
    if (coo != NULL)
	Win32OLEPropertySet(*coo, ST(1), ST(2));
}


SV *
FIRSTKEY(self)
    SV *self
PPCODE:
{
    SV **coo;
    unsigned int argErr;
    LPWINOLEOBJECT lpObj;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    BOOL bSuccess = FALSE;

    ST(0) = &sv_undef;
    coo = hv_fetch((HV*)SvRV(self), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DEB(fprintf(stderr, "Win32::OLE::Tie::FIRSTKEY |%s| |%d| |%lx|\n",
		PERL_OLE_ID, PERL_OLE_IDLEN,(long)coo));
    if (coo != NULL) {
	lpObj = (LPWINOLEOBJECT)SvIV(*coo);
	if (ValidDispatch(lpObj)) {
	    VariantInit(&result);

	    dispParams.rgvarg = NULL;
	    dispParams.rgdispidNamedArgs = NULL;
	    dispParams.cNamedArgs = 0;
	    dispParams.cArgs = 0;

	    EXCEPTIONINFOCLEAR(excepinfo);

	    LastOLEError = lpObj->pDispatch->Invoke(DISPID_NEWENUM, IID_NULL,
				lcidDefault,
				DISPATCH_METHOD | DISPATCH_PROPERTYGET,
				&dispParams, &result, &excepinfo, &argErr);

	    EXCEPTIONINFO(LastOLEError, excepinfo);
	    bSuccess = SUCCEEDED(LastOLEError);
	    if (bSuccess) {
		if ((result.vt&~VT_BYREF) == VT_UNKNOWN) {
		    IUnknown *punk;
		    IDispatch *pDisp;
		    IEnumVARIANT *pEnum;
		    if (result.vt&VT_BYREF) 
			punk = (IUnknown*)*result.ppunkVal;
		    else
			punk = (IUnknown*)result.punkVal;
		    if (SUCCEEDED(punk->QueryInterface(IID_IDispatch,
							(void**)&pDisp))) {
			CreatePerlObject((ST(0) = newSVsv(&sv_undef)),
					 pDisp,FALSE);
		    }
		    else if (SUCCEEDED(punk->QueryInterface(IID_IEnumVARIANT,
							    (void**)&pEnum))) {
			if (lpObj->pEnum != NULL)
			    lpObj->pEnum->Release();

			lpObj->pEnum = pEnum;
			VariantClear(&result);

			if (SUCCEEDED(pEnum->Reset())
			    && SUCCEEDED(pEnum->Next(1, &result, NULL)))
			    SetSVFromVariant(&result, (ST(0) = newSVpv("",0)));
		    }
		    punk->Release();
		}
		else
		    SetSVFromVariant(&result, (ST(0) = newSVpv("",0)));

		VariantClear(&result);
	    }
	}
	if (ST(0) != &sv_undef)
	    sv_2mortal(ST(0));
    }

    XSRETURN(1);
}

SV *
NEXTKEY(self,lastKey)
    SV *self
    SV *lastKey
PPCODE:
{
    SV **coo;
    LPWINOLEOBJECT lpObj;
    VARIANT result;
    BOOL bSuccess = FALSE;

    coo = hv_fetch((HV*)SvRV(ST(0)), PERL_OLE_ID, PERL_OLE_IDLEN, 0);
    DEB(fprintf(stderr, "Win32::OLE::Tie::NEXTKEY |%s| |%d| |%lx|\n",
		PERL_OLE_ID, PERL_OLE_IDLEN,(long)coo));
    if (coo != NULL) {
	lpObj = (LPWINOLEOBJECT)SvIV(*coo);
	if (ValidDispatch(lpObj) && lpObj->pEnum != NULL) {
	    VariantInit(&result);
	    bSuccess = lpObj->pEnum->Next(1, &result, NULL) == S_OK;
	    /* if the return is anything other than S_OK */
	    /* then this interface is no longer needed */
	    if (!bSuccess) {
		lpObj->pEnum->Release();
		lpObj->pEnum = NULL;
	    }
	}
    }

    if (bSuccess)
	SetSVFromVariant(&result, (ST(0) = sv_2mortal(newSVpv("",0))));
    else
	ST(0) = &sv_undef;

    XSRETURN(1);
}


