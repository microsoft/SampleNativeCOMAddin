/*!-----------------------------------------------------------------------
	connect.cpp

	The main implementation of the addin. It includes the implementation
	for IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer,
	and FormRegionStartup.
-----------------------------------------------------------------------!*/
#include "connect.h"
#include "formregionwrapper.h"

/*!-----------------------------------------------------------------------
	CConnect implementation
-----------------------------------------------------------------------!*/

_ATL_FUNC_INFO CConnect::DispatchFuncInfo = {CC_STDCALL, VT_EMPTY, 1, {VT_DISPATCH}};

CConnect::CConnect()
{
}

STDMETHODIMP CConnect::OnConnection
	( IDispatch *pApplication
	, ext_ConnectMode /* ConnectMode */
	, IDispatch* /* pAddInInst */
	, SAFEARRAY ** /* custom */ )
{
	if (!pApplication)
		return E_POINTER;

	m_pApplication = pApplication;

	ApplicationEventSink::DispEventAdvise(m_pApplication);

	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection
	( ext_DisconnectMode /*RemoveMode*/
	, SAFEARRAY ** /*custom*/ )
{
	if (m_pApplication)
		ApplicationEventSink::DispEventUnadvise(m_pApplication);

	m_pApplication.Release();
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate(SAFEARRAY ** /*custom*/)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete(SAFEARRAY ** /*custom*/)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown(SAFEARRAY ** /*custom*/)
{
	return S_OK;
}

STDMETHODIMP CConnect::Invoke
	( DISPID dispidMember
	, const IID &riid
	, LCID lcid
	, WORD wFlags
	, DISPPARAMS *pdispparams
	, VARIANT *pvarResult
	, EXCEPINFO *pexceptinfo
	, UINT *puArgErr )
{
	HRESULT hr;

	// Currently the CConnect object can get away with only one implementation
	// of Invoke because the only interfaces that Outlook calls Invoke on are
	// the ribbon callbacks and the form region startup. The other interfaces
	// IRibbonExtensibility, IDTExtensibility, and ICustomTaskPaneConsumer are
	// currently called via the virtual table and not the automation invoke
	// method although this could potentially change in the future. The key
	// thing to remember about using a common Invoke for multiple interfaces
	// is to ensure that dispids for the different interfaces don't overlap.

	// This is assuming the ribbon callback dispids are low and they not
	// intersect with any of the form region startup dispids
	hr = IRibbonCallbackImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

	if (DISP_E_MEMBERNOTFOUND == hr)
		hr = FormRegionStartupImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

	if (DISP_E_MEMBERNOTFOUND == hr)
		hr = IDTExtensibilityImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

	if (DISP_E_MEMBERNOTFOUND == hr)
		hr = ApplicationEventSink::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

	// There is no use trying to call Invoke on IRibbonExtensiblilityImpl and
	// ICustomTaskPaneConsumerImpl because they only contain one method with
	// dispid 1 and the other interfaces above most likely already cover dispid
	// 1, namely IDTExtensibility2, plus as already mention they are usually
	// called via the virtual table instead of IDispatch::Invoke.

	return hr;
}

/*!-----------------------------------------------------------------------
	FormRegionStartup interface implementation
-----------------------------------------------------------------------!*/

STDMETHODIMP CConnect::GetFormRegionStorage
	( BSTR /* bstrFormRegionName */
	, IDispatch * /* pDispItem */
	, long /* LCID */
	, Outlook::OlFormRegionMode /* formRegionMode */
	, Outlook::OlFormRegionSize /* formRegionSize */
	, __out VARIANT * pVarStorage )
{
	SAFEARRAY* pSafeArray = GetOFSResource(IDS_FORMREGIONSTORAGE);

	if (!pSafeArray)
		return E_UNEXPECTED;

	V_VT(pVarStorage) = VT_ARRAY | VT_UI1;
	V_ARRAY(pVarStorage) = pSafeArray;
	return S_OK;
}

STDMETHODIMP CConnect::BeforeFormRegionShow(_FormRegion *pFormRegion)
{
	return FormRegionWrapper::Setup(pFormRegion);
}

STDMETHODIMP CConnect::GetFormRegionManifest
	( BSTR /* bstrFormRegionName */
	, long /* LCID */
	, __out VARIANT * pvarManifest )
{
	BSTR bstr = GetXMLResource(IDS_FORMREGIONMANIFEST);

	if (!bstr)
		return E_UNEXPECTED;

	V_VT(pvarManifest) = VT_BSTR;
	V_BSTR(pvarManifest) = bstr;
	return S_OK;
}

STDMETHODIMP CConnect::GetFormRegionIcon
	( BSTR /* bstrFormRegionName */
	, long /* LCID */
	, Outlook::OlFormRegionIcon /* formRegionIcon */
	, __out VARIANT* /* pvarIcon */ )
{
	return S_OK;
}

/*!-----------------------------------------------------------------------
	ICustomTaskPaneConsumer interface implementation
-----------------------------------------------------------------------!*/

HRESULT CConnect::CTPFactoryAvailable(ICTPFactory *CTPFactoryInst)
{
	m_pCTPFactory = CTPFactoryInst;

	return HrCreateSampleTaskPane();
}

/*!-----------------------------------------------------------------------
	IRibbonExtensibility interface implementation
-----------------------------------------------------------------------!*/

HRESULT CConnect::GetCustomUI(BSTR /* ribbonID */, BSTR *ribbonXml)
{
	if(!ribbonXml)
		return E_POINTER;

	// Get the same ribbon xml for every ribbonID
	*ribbonXml = GetXMLResource(IDS_CUSTOMRIBBON);
	return S_OK;
}

HRESULT CConnect::Button1Clicked(IDispatch* /* ribbonControl */)
{
	MessageBoxW(NULL
		, L"Going to create a task pane now!"
		, L"Message from ribbon button."
		, MB_OK | MB_ICONINFORMATION);

	return HrCreateSampleTaskPane();
}

HRESULT CConnect::OptionsPagesAdd(IDispatch *pages)
{
	if (!pages)
		return E_POINTER;

	PropertyPagesPtr spPages(pages);

	if (!spPages)
		return E_UNEXPECTED;

	return spPages->Add(variant_t(SAMPLECONTROL_PROGID), bstr_t("Sample Options"));
}

HRESULT CConnect::HrCreateSampleTaskPane()
{
	if (!m_pCTPFactory)
		return E_POINTER;

	_CustomTaskPanePtr ctp;

	HRESULT hr = m_pCTPFactory->CreateCTP(bstr_t(SAMPLECONTROL_PROGID), bstr_t("Sample Task Pane"), vtMissing, &ctp);

	if (SUCCEEDED(hr))
		ctp->put_Visible(VARIANT_TRUE);

	return hr;
}

