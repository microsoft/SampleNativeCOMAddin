#include "TestMAPI.h"
#include "MAPIx.h"
#include "MAPIAux.h"
#include "Defaults.h"
#include <string>

#define INITGUID
#define USES_IID_IMAPIProp
#define USES_IID_IMessage
#include <guiddef.h>
#include "MAPIGuid.h"
DEFINE_GUID(IID_IMAPISecureMessage, 0x253cc320, 0xeab6, 0x11d0, 0x82, 0x22, 0, 0x60, 0x97, 0x93, 0x87, 0xea);

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

	void TestGetProp(LPMAPIPROP lpObj)
	{
		LPSPropValue lpProp = nullptr;
		auto cValues = 0UL;
		SPropTagArray sTag = { 1,{ PR_SUBJECT } };

		(void)lpObj->GetProps(
			&sTag,
			fMapiUnicode,
			&cValues,
			&lpProp);

		MAPIFreeBuffer(lpProp);
	}

	void TestBaseMessage(IMAPISecureMessage* lpObj)
	{
		LPMESSAGE base_item;
		auto hRes = lpObj->GetBaseMessage(&base_item);
		if (hRes == S_OK && base_item != nullptr)
		{
			//MessageBoxW(NULL, L"Got base message", L"TestBaseMessage", MB_OK | MB_ICONINFORMATION);
			TestGetProp(base_item);
			base_item->Release();
		}
	}

	void TestSecureMessage(LPMAPIPROP lpObj)
	{
		IMAPISecureMessage* secure_message = nullptr;
		auto hRes = lpObj->QueryInterface(IID_IMAPISecureMessage, (LPVOID*)&secure_message);
		if (hRes == S_OK && nullptr != secure_message)
		{
			//MessageBoxW(NULL, L"Got secure message", L"TestSecureMessage", MB_OK | MB_ICONINFORMATION);
			TestBaseMessage(secure_message);
			TestGetProp(lpObj);

			secure_message->Release();
		}

		TestGetProp(lpObj);
	}
}
