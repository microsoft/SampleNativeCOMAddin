/*!-----------------------------------------------------------------------
	samplecontrol.cpp
-----------------------------------------------------------------------!*/
#include "stdafx.h"
#include "SampleControl.h"


/*!-----------------------------------------------------------------------
	CSampleControl implementation
-----------------------------------------------------------------------!*/

LRESULT CSampleControl::OnBnClickedButton1(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CComPtr<IOleClientSite> spClientSite;
	if (SUCCEEDED(GetClientSite(&spClientSite)))
	{
		PropertyPageSitePtr propertyPageSite(spClientSite.p);

		if (propertyPageSite)
		{
			propertyPageSite->OnStatusChange();
		}
	}

	::MessageBoxW(nullptr,
		L"You clicked button1 on our SampleControl!",
		L"Message from sample control",
		MB_OK | MB_ICONINFORMATION);

	return S_OK;
}

HRESULT CSampleControl::get_Dirty(VARIANT_BOOL *dirty)
{
	if (!dirty)
		return E_POINTER;

	*dirty = VARIANT_TRUE;
	return S_OK;
}

HRESULT CSampleControl::Apply()
{
	// Save any settings from the property page if needed.
	return S_OK;
}