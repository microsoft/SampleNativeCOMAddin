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

	const auto hModule = _AtlBaseModule.GetModuleInstance();

	if (!hModule)
		return E_UNEXPECTED;

	const auto hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);

	if (!hRsrc)
		return HRESULT_FROM_WIN32(GetLastError());

	const auto hGlobal = LoadResource(hModule, hRsrc);

	if (!hGlobal)
		return HRESULT_FROM_WIN32(GetLastError());

	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);

	return S_OK;
}

BSTR GetXMLResource(int nId)
{
	LPVOID pResourceData = nullptr;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("XML"), &pResourceData, &dwSizeInBytes)))
		return nullptr;

	// Assumes that the data is not stored in Unicode.
	CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));

	return cbstr.Detach();
}

SAFEARRAY* GetOFSResource(int nId)
{
	LPVOID pResourceData = nullptr;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("OFS"), &pResourceData, &dwSizeInBytes)))
		return nullptr;

	SAFEARRAY* psa = nullptr;
	SAFEARRAYBOUND dim = { dwSizeInBytes, 0 };

	psa = SafeArrayCreate(VT_UI1, 1, &dim);

	if (!psa)
		return nullptr;

	BYTE* pSafeArrayData = nullptr;

	if (FAILED(SafeArrayAccessData(psa, reinterpret_cast<void**>(&pSafeArrayData))))
	{
		SafeArrayDestroy(psa);
		return nullptr;
	}

	memcpy(pSafeArrayData, pResourceData, dwSizeInBytes);

	if (FAILED(SafeArrayUnaccessData(psa)))
	{
		SafeArrayDestroy(psa);
		return nullptr;
	}

	return psa;
}
