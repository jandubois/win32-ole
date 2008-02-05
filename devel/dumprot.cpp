// Debuggin tool: dump Running-Object-Table
// compile with MSC++ "cl -W3 dumprot.cpp ole32.lib"

#include <stdio.h>
#include <windows.h>

int main()
{
    LPRUNNINGOBJECTTABLE pROT = NULL;
    LPMALLOC pMalloc = NULL;

    if (FAILED(CoGetMalloc(1,&pMalloc))) {
	printf("CoGetMalloc failed!\n");
	return 1;
    }

    if (FAILED(CoInitialize(NULL))) {
	printf("CoInitialize failed!\n");
	return 1;
    }

    if (GetRunningObjectTable(0, &pROT) != S_OK)
	printf("GetRunningObjectTable failed!\n");
    else {
	IEnumMoniker *pEnumMoniker = NULL;

	if (FAILED(pROT->EnumRunning(&pEnumMoniker)))
	    printf("EnumRunning failed!\n");
	else {
	    ULONG ulCount;
	    IMoniker *pMoniker;
	    IBindCtx *pBindCtx;

	    CreateBindCtx(0, &pBindCtx);
	    while (SUCCEEDED(pEnumMoniker->Next(1, &pMoniker, &ulCount))
		   && ulCount > 0)
	    {
		OLECHAR *pDisplayName = NULL;

		if (SUCCEEDED(pMoniker->GetDisplayName(pBindCtx, NULL,
						       &pDisplayName)))
		{
		    wprintf(L"Moniker is \"%s\"\n", pDisplayName);
		    pMalloc->Free(pDisplayName);
		}
		pMoniker->Release();
	    }
	    pBindCtx->Release();
	    pEnumMoniker->Release();
	}
	pROT->Release();
    }
    CoUninitialize();
    return 0;
}
