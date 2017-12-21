#include "ApplicationEventsSink.h"
#include "MAPIx.h"
#include "MAPI\TestMAPI.h"

_ATL_FUNC_INFO ApplicationEventsSink::OptionsPagesAddInfo = { CC_STDCALL, VT_EMPTY, 1,{ VT_DISPATCH } };
_ATL_FUNC_INFO ApplicationEventsSink::MapiLogonCompleteInfo = { CC_STDCALL, VT_EMPTY, 0 };
_ATL_FUNC_INFO ApplicationEventsSink::ItemSendInfo = { CC_STDCALL, VT_EMPTY, 2,{ VT_DISPATCH, VT_BOOL | VT_BYREF } };

ApplicationEventsSink::ApplicationEventsSink(Outlook::_ApplicationPtr piApp)
{
	m_piApp = piApp;
	DispEventAdvise((IUnknown*)m_piApp);
}

ApplicationEventsSink::~ApplicationEventsSink()
{
	DispEventUnadvise((IUnknown*)m_piApp);
}

HRESULT ApplicationEventsSink::OptionsPagesAdd(IDispatch *pages)
{
	if (!pages)
		return E_POINTER;

	PropertyPagesPtr spPages(pages);

	if (!spPages)
		return E_UNEXPECTED;

	return spPages->Add(variant_t(SAMPLECONTROL_PROGID), bstr_t("Sample Options"));
}

HRESULT ApplicationEventsSink::MapiLogonComplete()
{
	//MessageBoxW(NULL, L"MapiLogonComplete", L"Sample Add-In", MB_OK | MB_ICONINFORMATION);
	return S_OK;
}

HRESULT ApplicationEventsSink::ItemSend(IDispatch* /*Item*/, VARIANT_BOOL* /*Cancel*/)
{
	//MessageBoxW(NULL, L"ItemSend", L"Sample Add-In", MB_OK | MB_ICONINFORMATION);
	//Outlook::_MailItemPtr oom_item = Item;
	//LPUNKNOWN mapi_object = nullptr;
	//oom_item->get_MAPIOBJECT(&mapi_object);
	//if (mapi_object != nullptr)
	//{
	//	LPMAPIPROP mapi_item = nullptr;
	//	mapi_object->QueryInterface(IID_IMAPIProp, (LPVOID*)&mapi_item);
	//	if (mapi_item != nullptr)
	//	{
	//		TestMAPI::TestSecureMessage(mapi_item);

	//		mapi_item->Release();
	//	}

	//	mapi_object->Release();
	//}

	return S_OK;
}
