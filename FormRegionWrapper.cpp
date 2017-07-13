/*!-----------------------------------------------------------------------
	formregionwrapper.cpp
-----------------------------------------------------------------------!*/
#include "stdafx.h"
#include "formregionwrapper.h"

#define ReturnHrOnFailure(h) { hr = (h); ATLASSERT(SUCCEEDED((hr))); if (FAILED(hr)) return hr; }

/*!-----------------------------------------------------------------------
	FormRegionWrapper implementation
-----------------------------------------------------------------------!*/

_ATL_FUNC_INFO FormRegionWrapper::VoidFuncInfo = { CC_STDCALL, VT_EMPTY, 0, 0 };

/*static*/ HRESULT FormRegionWrapper::Setup(_FormRegion* pFormRegion) throw()
{
	if (!pFormRegion)
		return E_POINTER;

	// The FormRegionWrapper object is created here and
	// deleted in the FormRegionClose event
	FormRegionWrapper* pWrapper = new (std::nothrow) FormRegionWrapper();

	if (!pWrapper)
		return E_OUTOFMEMORY;

	return pWrapper->HrInit(pFormRegion);
}

HRESULT FormRegionWrapper::HrInit(_FormRegion* pFormRegion)
{
	HRESULT hr = S_OK;

	if (!pFormRegion)
		return E_POINTER;

	m_spFormRegion = pFormRegion;

	// Subscribe to the form region events
	FormRegionEventSink::DispEventAdvise(m_spFormRegion);

	// Get the form so we can subscribe to the button click event
	// of the command button on the form region
	IDispatchPtr spDispatch;
	ReturnHrOnFailure(pFormRegion->get_Form(&spDispatch));

	_UserFormPtr spForm;
	ReturnHrOnFailure(spDispatch->QueryInterface(&spForm));

	ControlsPtr spControls;
	ReturnHrOnFailure(spForm->get_Controls(&spControls));

	IControlPtr spControl;
	ReturnHrOnFailure(spControls->_GetItemByName(bstr_t("OlkCommandButton1"), &spControl));

	ReturnHrOnFailure(spControl->QueryInterface(&m_spOlkCmdBtn));

	// Subscribe to the button events on the form region
	OlkCommandButtonEventSink::DispEventAdvise(m_spOlkCmdBtn);

	// Store the item that this form region is connected to
	IDispatchPtr spDispItem;
	ReturnHrOnFailure(pFormRegion->get_Item(&spDispItem));

	ReturnHrOnFailure(spDispItem->QueryInterface(&m_spMailItem));

	return hr;
}

void FormRegionWrapper::OnButton1Click()
{
	// When the button on the form region is clicked send the item
	if (m_spMailItem)
	{
		MessageBoxW(NULL,
			L"Going to send the message now!",
			L"Message from the sample native form region",
			MB_OK | MB_ICONINFORMATION);

		m_spMailItem->Send();

		// After sending don't access any members of the FormRegionWrapper
		// because it is likely destroyed because the OnFormRegionClose
		// should have been called and it deletes the object.
	}
}

void FormRegionWrapper::OnFormRegionClose()
{
	// Clean up the button reference and event sink
	if (m_spOlkCmdBtn)
	{
		OlkCommandButtonEventSink::DispEventUnadvise(m_spOlkCmdBtn);
		m_spOlkCmdBtn.Release();
	}

	// Clean up the form region reference and event sink
	if (m_spFormRegion)
	{
		FormRegionEventSink::DispEventUnadvise(m_spFormRegion);
		m_spFormRegion.Release();
	}

	m_spMailItem.Release();

	// Clean up this object
	delete this;
}
