#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>

/*
 *  MAPI Stub Utilities
 *
 *	Public Functions:
 *
 *		GetPrivateMAPI()
 *			Obtain a handle to the MAPI DLL.  This function will load the MAPI DLL
 *			if it hasn't already been loaded
 *
 *		UnLoadPrivateMAPI()
 *			Forces the MAPI DLL to be unloaded.  This can cause problems if the code
 *			still has outstanding allocated MAPI memory, or unmatched calls to
 *			MAPIInitialize/MAPIUninitialize
 *
 *		ForceOutlookMAPI()
 *			Instructs the stub code to always try loading the Outlook version of MAPI
 *			on the system, instead of respecting the system MAPI registration
 *			(HKLM\Software\Clients\Mail). This call must be made prior to any MAPI
 *			function calls.
 */
HMODULE GetPrivateMAPI();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI(bool fForce);

const WCHAR WszKeyNameMailClient[] = L"Software\\Clients\\Mail";
const WCHAR WszValueNameDllPathEx[] = L"DllPathEx";
const WCHAR WszValueNameDllPath[] = L"DllPath";

const CHAR SzValueNameMSI[] = "MSIComponentID";
const CHAR SzValueNameLCID[] = "MSIApplicationLCID";

const WCHAR WszOutlookMapiClientName[] = L"Microsoft Outlook";

const WCHAR WszMAPISystemPath[] = L"%s\\%s";
const WCHAR WszMAPISystemDrivePath[] = L"%s%s%s";
const WCHAR szMAPISystemDrivePath[] = L"%hs%hs%ws";

static const WCHAR WszOlMAPI32DLL[] = L"olmapi32.dll";
static const WCHAR WszMSMAPI32DLL[] = L"msmapi32.dll";
static const WCHAR WszMapi32[] = L"mapi32.dll";
static const WCHAR WszMapiStub[] = L"mapistub.dll";

static const CHAR SzFGetComponentPath[] = "FGetComponentPath";

// Sequence number which is incremented every time we set our MAPI handle which will
//  cause a re-fetch of all stored function pointers
volatile ULONG g_ulDllSequenceNum = 1;

// Whether or not we should ignore the system MAPI registration and always try to find
//  Outlook and its MAPI DLLs
static bool s_fForceOutlookMAPI = false;

// Whether or not we should ignore the registry and load MAPI from the system directory
static bool s_fForceSystemMAPI = false;

static volatile HMODULE g_hinstMAPI = NULL;
HMODULE g_hModPstPrx32 = NULL;

__inline HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
}

enum mapiSource;
class MAPIPathIterator
{
public:
	MAPIPathIterator(bool bBypassRestrictions);
	~MAPIPathIterator();
	LPWSTR GetNextMAPIPath();
	LPWSTR GetMAPISystemDir();

private:
	LPWSTR GetRegisteredMapiClient(LPCWSTR pwzProviderOverride, bool bDLL, bool bEx);
	LPWSTR GetMailClientFromMSIData(HKEY hkeyMapiClient);
	LPWSTR GetMailClientFromDllPath(HKEY hkeyMapiClient, bool bEx);

	mapiSource CurrentSource;
	HKEY m_hMailKey;
	HKEY m_hkeyMapiClient;
	LPWSTR m_rgchMailClient;
	LPCWSTR m_szRegisteredClient;
	bool m_bBypassRestrictions;

	int m_iCurrentOutlook;
};

// Keep this in sync with g_pszOutlookQualifiedComponents
#define oqcOfficeBegin   0
#define oqcOffice15      oqcOfficeBegin + 0
#define oqcOffice14      oqcOfficeBegin + 1
#define oqcOffice12      oqcOfficeBegin + 2
#define oqcOffice11      oqcOfficeBegin + 3
#define oqcOffice11Debug oqcOfficeBegin + 4
#define oqcOfficeEnd     oqcOffice11Debug

void SetMAPIHandle(HMODULE hinstMAPI)
{
	HMODULE	hinstNULL = NULL;
	HMODULE	hinstToFree = NULL;

	if (hinstMAPI == NULL)
	{
		// If we've preloaded pstprx32.dll, unload it before MAPI is unloaded to prevent dependency problems
		if (g_hModPstPrx32)
		{
			::FreeLibrary(g_hModPstPrx32);
			g_hModPstPrx32 = NULL;
		}

		hinstToFree = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstNULL);
	}
	else
	{
		// Set the value only if the global is NULL
		HMODULE	hinstPrev;
		// Code Analysis gives us a C28112 error when we use InterlockedCompareExchangePointer, so we instead exchange, check and exchange back
		//hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
		hinstPrev = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstMAPI);
		if (NULL != hinstPrev)
		{
			(void)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstPrev);
			hinstToFree = hinstMAPI;
		}

		// If we've updated our MAPI handle, any previous addressed fetched via GetProcAddress are invalid, so we
		// have to increment a sequence number to signal that they need to be re-fetched
		InterlockedIncrement(reinterpret_cast<volatile LONG*>(&g_ulDllSequenceNum));
	}
	if (NULL != hinstToFree)
	{
		FreeLibrary(hinstToFree);
	}
}

/*
 *  RegQueryWszExpand
 *		Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
DWORD RegQueryWszExpand(HKEY hKey, LPCWSTR lpValueName, LPWSTR lpValue, DWORD cchValueLen)
{
	DWORD dwErr = ERROR_SUCCESS;
	DWORD dwType = 0;

	WCHAR rgchValue[MAX_PATH] = { 0 };
	DWORD dwSize = sizeof(rgchValue);

	dwErr = RegQueryValueExW(hKey, lpValueName, 0, &dwType, (LPBYTE)&rgchValue, &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		if (dwType == REG_EXPAND_SZ)
		{
			// Expand the strings
			DWORD cch = ExpandEnvironmentStringsW(rgchValue, lpValue, cchValueLen);
			if ((0 == cch) || (cch > cchValueLen))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
		}
		else if (dwType == REG_SZ)
		{
			wcscpy_s(lpValue, cchValueLen, rgchValue);
		}
	}
Exit:
	return dwErr;
}

/*
 *  GetComponentPath
 *		Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to
 *		a DLL location from the default MAPI client registration values
 */
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall)
{
	HMODULE hMapiStub = NULL;
	bool fReturn = FALSE;

	typedef bool (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

	hMapiStub = LoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = LoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		FGetComponentPathType pFGetCompPath = (FGetComponentPathType)GetProcAddress(hMapiStub, SzFGetComponentPath);

		if (pFGetCompPath)
		{
			fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);
		}

		FreeLibrary(hMapiStub);
	}

	return fReturn;
} // GetComponentPath

enum mapiSource
{
	msRegisteredMSI,
	msRegisteredDLLEx,
	msRegisteredDLL,
	msSystem,
	msEnd,
};

MAPIPathIterator::MAPIPathIterator(bool bBypassRestrictions)
{
	m_bBypassRestrictions = bBypassRestrictions;
	m_szRegisteredClient = NULL;
	if (bBypassRestrictions)
	{
		CurrentSource = msRegisteredMSI;
	}
	else
	{
		if (!s_fForceSystemMAPI)
		{
			CurrentSource = msRegisteredMSI;
			if (s_fForceOutlookMAPI)
				m_szRegisteredClient = WszOutlookMapiClientName;
		}
		else
			CurrentSource = msSystem;
	}
	m_hMailKey = NULL;
	m_hkeyMapiClient = NULL;
	m_rgchMailClient = NULL;

	m_iCurrentOutlook = oqcOfficeBegin;
}

MAPIPathIterator::~MAPIPathIterator()
{
	delete[] m_rgchMailClient;
	if (m_hMailKey) RegCloseKey(m_hMailKey);
	if (m_hkeyMapiClient) RegCloseKey(m_hkeyMapiClient);
}

LPWSTR MAPIPathIterator::GetNextMAPIPath()
{
	// Mini state machine here will get the path from the current source then set the next source to search
	// Either returns the next available MAPI path or NULL if none remain
	LPWSTR szPath = NULL;
	while (msEnd != CurrentSource && !szPath)
	{
		switch (CurrentSource)
		{
		case msRegisteredMSI:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, false, false);
			CurrentSource = msRegisteredDLLEx;
			break;
		case msRegisteredDLLEx:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, true);
			CurrentSource = msRegisteredDLL;
			break;
		case msRegisteredDLL:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, false);
			if (s_fForceOutlookMAPI && !m_bBypassRestrictions)
			{
				CurrentSource = msEnd;
			}
			else
			{
				CurrentSource = msSystem;
			}
			break;
		case msSystem:
			szPath = GetMAPISystemDir();
			CurrentSource = msEnd;
			break;
		case msEnd:
		default:
			break;
		}
	}
	return szPath;
}

// if cchszA == -1, MultiByteToWideChar will compute the length
// Delete with delete[]
_Check_return_ void AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA)
{
	if (!ppszW) return;
	*ppszW = NULL;
	if (NULL == pszA) return;
	if (!cchszA) return;

	// Get our buffer size
	int iRet = 0;
	iRet = MultiByteToWideChar(
		CP_ACP,
		0,
		pszA,
		(int)cchszA,
		NULL,
		NULL);
	if (0 != iRet)
	{
		// MultiByteToWideChar returns num of chars
		LPWSTR pszW = new WCHAR[iRet];

		iRet = MultiByteToWideChar(
			CP_ACP,
			0,
			pszA,
			(int)cchszA,
			pszW,
			iRet);
		if (0 != iRet)
		{
			*ppszW = pszW;
		}
		else
		{
			delete[] pszW;
		}
	}
}

/*
 *  GetMailClientFromMSIData
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
LPWSTR MAPIPathIterator::GetMailClientFromMSIData(HKEY hkeyMapiClient)
{
	CHAR rgchMSIComponentID[MAX_PATH] = { 0 };
	CHAR rgchMSIApplicationLCID[MAX_PATH] = { 0 };
	CHAR rgchComponentPath[MAX_PATH] = { 0 };
	DWORD dwType = 0;
	LPWSTR szPath = NULL;

	DWORD dwSizeComponentID = sizeof(rgchMSIComponentID);
	DWORD dwSizeLCID = sizeof(rgchMSIApplicationLCID);

	if (ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameMSI, 0, &dwType, (LPBYTE)&rgchMSIComponentID, &dwSizeComponentID) &&
		ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameLCID, 0, &dwType, (LPBYTE)&rgchMSIApplicationLCID, &dwSizeLCID))
	{
		if (GetComponentPath(rgchMSIComponentID, rgchMSIApplicationLCID, rgchComponentPath, _countof(rgchComponentPath), FALSE))
		{
			AnsiToUnicode(rgchComponentPath, &szPath, (size_t)-1);
		}
	}

	return szPath;
}

/*
 *  GetMailClientFromDllPath
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
LPWSTR MAPIPathIterator::GetMailClientFromDllPath(HKEY hkeyMapiClient, bool bEx)
{
	LPWSTR szPath = NULL;
	DWORD ret = ERROR_SUCCESS;

	szPath = new WCHAR[MAX_PATH];

	if (szPath)
	{
		if (bEx)
		{
			ret = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPathEx, szPath, MAX_PATH);
		}
		else
		{
			ret = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPath, szPath, MAX_PATH);
		}
		if (ERROR_SUCCESS == ret)
		{
			delete[] szPath;
			szPath = NULL;
		}
	}

	return szPath;
}

/*
 *  GetRegisteredMapiClient
 *		Read the registry to discover the registered MAPI client and attempt to load its MAPI DLL.
 *
 *		If wzOverrideProvider is specified, this function will load that MAPI Provider instead of the
 *		currently registered provider
 */
LPWSTR MAPIPathIterator::GetRegisteredMapiClient(LPCWSTR pwzProviderOverride, bool bDLL, bool bEx)
{
	DWORD ret = ERROR_SUCCESS;
	LPWSTR szPath = NULL;
	LPCWSTR pwzProvider = pwzProviderOverride;

	if (!m_hMailKey)
	{
		// Open HKLM\Software\Clients\Mail
		ret = RegOpenKeyExW(HKEY_LOCAL_MACHINE,
			WszKeyNameMailClient,
			0,
			KEY_READ,
			&m_hMailKey);
		if (ERROR_SUCCESS == ret)
		{
			m_hMailKey = NULL;
		}
	}

	// If a specific provider wasn't specified, load the name of the default MAPI provider
	if (m_hMailKey && !pwzProvider && !m_rgchMailClient)
	{
		m_rgchMailClient = new WCHAR[MAX_PATH];
		if (m_rgchMailClient)
		{
			// Get Outlook application path registry value
			DWORD dwSize = MAX_PATH;
			DWORD dwType = 0;
			ret = RegQueryValueExW(
				m_hMailKey,
				NULL,
				0,
				&dwType,
				(LPBYTE)m_rgchMailClient,
				&dwSize);
			if (ERROR_SUCCESS != ret)
			{
				delete[] m_rgchMailClient;
				m_rgchMailClient = NULL;
			}
		}
	}

	if (!pwzProvider) pwzProvider = m_rgchMailClient;

	if (m_hMailKey && pwzProvider && !m_hkeyMapiClient)
	{
		ret = RegOpenKeyExW(
			m_hMailKey,
			pwzProvider,
			0,
			KEY_READ,
			&m_hkeyMapiClient);
		if (ERROR_SUCCESS != ret)
		{
			m_hkeyMapiClient = NULL;
		}
	}

	if (m_hkeyMapiClient)
	{
		if (bDLL)
		{
			szPath = GetMailClientFromDllPath(m_hkeyMapiClient, bEx);
		}
		else
		{
			szPath = GetMailClientFromMSIData(m_hkeyMapiClient);
		}
	}

	return szPath;
}

/*
 *  GetMAPISystemDir
 *		Fall back for loading System32\Mapi32.dll if all else fails
 */
LPWSTR MAPIPathIterator::GetMAPISystemDir()
{
	WCHAR szSystemDir[MAX_PATH] = { 0 };

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		LPWSTR szDLLPath = new WCHAR[MAX_PATH];
		if (szDLLPath)
		{
			swprintf_s(szDLLPath, MAX_PATH, WszMAPISystemPath, szSystemDir, WszMapi32);
			return szDLLPath;
		}
	}

	return NULL;
}

WCHAR g_pszOutlookQualifiedComponents[][MAX_PATH] = {
	L"{E83B4360-C208-4325-9504-0D23003A74A5}", // O15_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	L"{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}", // O14_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	L"{24AAE126-0911-478F-A019-07B875EB9996}", // O12_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	L"{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	L"{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (debug)  // STRING_OK
};

HMODULE GetDefaultMapiHandle()
{
	HMODULE hinstMapi = NULL;

	LPWSTR szPath = NULL;
	MAPIPathIterator* mpi = new MAPIPathIterator(false);

	if (mpi)
	{
		while (!hinstMapi)
		{
			szPath = mpi->GetNextMAPIPath();
			if (!szPath) break;

			hinstMapi = LoadLibraryW(szPath);
			delete[] szPath;
		}
	}

	delete mpi;
	return hinstMapi;
}

/*------------------------------------------------------------------------------
	Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	current process.
	------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	HMODULE	hinstPrivateMAPI = NULL;
	GetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	return hinstPrivateMAPI;
}

void UnLoadPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = NULL;

	hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(NULL);
	}
}

void ForceOutlookMAPI(bool fForce)
{
	s_fForceOutlookMAPI = fForce;
}

void ForceSystemMAPI(bool fForce)
{
	s_fForceSystemMAPI = fForce;
}

HMODULE GetPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = GetMAPIHandle();

	if (NULL == hinstPrivateMAPI)
	{
		// First, try to attach to olmapi32.dll if it's loaded in the process
		hinstPrivateMAPI = AttachToMAPIDll(WszOlMAPI32DLL);

		// If that fails try msmapi32.dll, for Outlook 11 and below
		//  Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = AttachToMAPIDll(WszMSMAPI32DLL);
		}

		// If MAPI isn't loaded in the process yet, then find the path to the DLL and
		// load it manually.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = GetDefaultMapiHandle();
		}

		if (NULL != hinstPrivateMAPI)
		{
			SetMAPIHandle(hinstPrivateMAPI);
		}

		// Reason - if for any reason there is an instance already loaded, SetMAPIHandle()
		// will free the new one and reuse the old one
		// So we fetch the instance from the global again
		return GetMAPIHandle();
	}

	return hinstPrivateMAPI;
}