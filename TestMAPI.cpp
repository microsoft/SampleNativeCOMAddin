#include "TestMAPI.h"
#include "MAPIx.h"
#include "MAPIAux.h"
#include "MAPI\Defaults.h"
#include <string>

namespace TestMAPI {
	void TestInbox(std::wstring caller, bool bOnline)
	{
		LPMAPISESSION lpMAPISession = NULL;

		auto hRes = MAPILogonEx(0, NULL, NULL,
			MAPI_LOGON_UI,
			&lpMAPISession);
		if (SUCCEEDED(hRes))
		{
			LPMDB lpMDB = NULL;
			hRes = OpenDefaultMessageStore(lpMAPISession, bOnline ? MDB_ONLINE : NULL, &lpMDB);
			if (SUCCEEDED(hRes))
			{
				LPMAPIFOLDER lpInbox = NULL;
				hRes = OpenInbox(lpMDB, bOnline ? MAPI_NO_CACHE : NULL, &lpInbox);
				if (SUCCEEDED(hRes))
				{
					auto message = caller + L": got Inbox " + (bOnline ? L"online" : L"cached");
					MessageBoxW(NULL, message.c_str(), L"Sample Add-In", MB_OK | MB_ICONINFORMATION);
				}

				if (lpInbox) lpInbox->Release();
				lpInbox = NULL;
			}

			if (lpMDB) lpMDB->Release();
			lpMDB = NULL;
		}

		if (lpMAPISession) lpMAPISession->Release();
		lpMAPISession = NULL;
	}
}
