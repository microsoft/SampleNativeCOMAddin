/*!-----------------------------------------------------------------------
	resource.cpp
-----------------------------------------------------------------------!*/
#include "stdafx.h"

/*!-----------------------------------------------------------------------
	Retrieves a resource from the module.

	According to MSDN there is no need to clean up the resources
	because they will be released when the module is unloaded.
-----------------------------------------------------------------------!*/
HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
	if (!lpType || !ppvResourceData || !pdwSizeInBytes)
		return E_POINTER;

	HMODULE hModule = _AtlBaseModule.GetModuleInstance();

	if (!hModule)
		return E_UNEXPECTED;

	HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);

	if (!hRsrc)
		return HRESULT_FROM_WIN32(GetLastError());

	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);

	if (!hGlobal)
		return HRESULT_FROM_WIN32(GetLastError());

	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);

	return S_OK;
}

BSTR GetXMLResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("XML"), &pResourceData, &dwSizeInBytes)))
		return NULL;

	// Assumes that the data is not stored in Unicode.
	CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));

	return cbstr.Detach();
}

SAFEARRAY* GetOFSResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("OFS"), &pResourceData, &dwSizeInBytes)))
		return NULL;

	SAFEARRAY* psa = NULL;
	SAFEARRAYBOUND dim = {dwSizeInBytes, 0};

	psa = SafeArrayCreate(VT_UI1, 1, &dim);

	if (!psa)
		return NULL;

	BYTE* pSafeArrayData = NULL;

	if (FAILED(SafeArrayAccessData(psa, (void**)&pSafeArrayData)))
	{
		SafeArrayDestroy(psa);
		return NULL;
	}

	memcpy(pSafeArrayData, pResourceData, dwSizeInBytes);

	if (FAILED(SafeArrayUnaccessData(psa)))
	{
		SafeArrayDestroy(psa);
		return NULL;
	}

	return psa;
}
