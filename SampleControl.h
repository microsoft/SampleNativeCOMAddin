/*!-----------------------------------------------------------------------
	samplecontrol.h
-----------------------------------------------------------------------!*/
#pragma once
#include "stdafx.h"
#include <atlctl.h>
#include "outlook.h"

typedef
IDispatchImpl<PropertyPage, &__uuidof(PropertyPage), &__uuidof(__Outlook), -1, -1>
PropertyPageImpl;

// CSampleControl
class ATL_NO_VTABLE CSampleControl
	: public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<ISampleControl, &__uuidof(ISampleControl), &__uuidof(__SampleNativeCOMAddinLib), /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IPersistStreamInitImpl<CSampleControl>,
	public IOleControlImpl<CSampleControl>,
	public IOleObjectImpl<CSampleControl>,
	public IOleInPlaceActiveObjectImpl<CSampleControl>,
	public IViewObjectExImpl<CSampleControl>,
	public IOleInPlaceObjectWindowlessImpl<CSampleControl>,
	public ISupportErrorInfo,
	public IPersistStorageImpl<CSampleControl>,
	public ISpecifyPropertyPagesImpl<CSampleControl>,
	public IQuickActivateImpl<CSampleControl>,
	public IProvideClassInfo2Impl<&__uuidof(SampleControl), NULL, &__uuidof(__SampleNativeCOMAddinLib)>,
	public CComCoClass<CSampleControl, &__uuidof(SampleControl)>,
	public CComCompositeControl<CSampleControl>,
	public PropertyPageImpl
{
public:

	CSampleControl()
	{
		m_bWindowOnly = TRUE;
		CalcExtent(m_sizeExtent);
	}

	DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
	OLEMISC_CANTLINKINSIDE |
		OLEMISC_INSIDEOUT |
		OLEMISC_ACTIVATEWHENVISIBLE |
		OLEMISC_SETCLIENTSITEFIRST
		)

		static HRESULT WINAPI UpdateRegistry(BOOL bRegister) throw()
	{
		ATL::_ATL_REGMAP_ENTRY regMapEntries[] =
		{
			{ OLESTR("OLEMISC"), NULL },
			{ OLESTR("PROGID"), SAMPLECONTROL_PROGID },
			{ OLESTR("CLSID"), SAMPLECONTROL_CLSID_STR },
			{ OLESTR("TYPELIB"), TYPELIB_GUID_STR },
			{ NULL, NULL }
		};

		TCHAR szOleMisc[32];
		ATL::Checked::itot_s(_GetMiscStatus(), szOleMisc, _countof(szOleMisc), 10);
		USES_CONVERSION_EX;
		regMapEntries[0].szData = T2OLE_EX(szOleMisc, _ATL_SAFE_ALLOCA_DEF_THRESHOLD);

		return ATL::_pAtlModule->UpdateRegistryFromResource(IDR_SAMPLECONTROL, bRegister, regMapEntries);
	}


	BEGIN_COM_MAP(CSampleControl)
		COM_INTERFACE_ENTRY(ISampleControl)
		COM_INTERFACE_ENTRY2(IDispatch, PropertyPage)
		COM_INTERFACE_ENTRY(IViewObjectEx)
		COM_INTERFACE_ENTRY(IViewObject2)
		COM_INTERFACE_ENTRY(IViewObject)
		COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceObject)
		COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
		COM_INTERFACE_ENTRY(IOleControl)
		COM_INTERFACE_ENTRY(IOleObject)
		COM_INTERFACE_ENTRY(IPersistStreamInit)
		COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
		COM_INTERFACE_ENTRY(IQuickActivate)
		COM_INTERFACE_ENTRY(IPersistStorage)
		COM_INTERFACE_ENTRY(IProvideClassInfo)
		COM_INTERFACE_ENTRY(IProvideClassInfo2)
		COM_INTERFACE_ENTRY(PropertyPage)
	END_COM_MAP()

	BEGIN_PROP_MAP(CSampleControl)
		PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
		PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	END_PROP_MAP()


	BEGIN_MSG_MAP(CSampleControl)
		COMMAND_HANDLER(IDC_BUTTON1, BN_CLICKED, OnBnClickedButton1)
		CHAIN_MSG_MAP(CComCompositeControl<CSampleControl>)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CSampleControl)
		//Make sure the Event Handlers have __stdcall calling convention
	END_SINK_MAP()

	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
	{
		if (dispid == DISPID_AMBIENT_BACKCOLOR)
		{
			SetBackgroundColorFromAmbient();
			FireViewChange();
		}
		return IOleControlImpl<CSampleControl>::OnAmbientPropertyChange(dispid);
	}

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
	{
		static const IID* arr[] =
		{
			&__uuidof(ISampleControl),
		};

		for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++)
		{
			if (InlineIsEqualGUID(*arr[i], riid))
				return S_OK;
		}
		return S_FALSE;
	}

	// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

	// ISampleControl

	enum { IDD = IDD_SAMPLECONTROL };

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct() { return S_OK; }
	void FinalRelease() { }

	LRESULT OnBnClickedButton1(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	// PropertyPage
	STDMETHOD(GetPageInfo)(BSTR*, LONG*)
	{
		// This method is never used on the PropertyPage interface
		return E_NOTIMPL;
	}
	STDMETHOD(get_Dirty)(VARIANT_BOOL* dirty);
	STDMETHOD(Apply)();
};

OBJECT_ENTRY_AUTO(__uuidof(SampleControl), CSampleControl)
