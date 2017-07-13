/*!-----------------------------------------------------------------------
	connect.h
-----------------------------------------------------------------------!*/
#pragma once
#include "stdafx.h"

class CConnect;

typedef
IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
IDTExtensibilityImpl;

typedef
IDispatchImpl<_FormRegionStartup, &__uuidof(_FormRegionStartup), &__uuidof(__Outlook), -1, -1>
FormRegionStartupImpl;

typedef
IDispatchImpl<ICustomTaskPaneConsumer, &__uuidof(ICustomTaskPaneConsumer), &__uuidof(__Office), -1, -1>
ICustomTaskPaneConsumerImpl;

typedef
IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), -1, -1>
IRibbonExtensibilityImpl;

typedef
IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback), &__uuidof(__SampleNativeCOMAddinLib), -1, -1>
IRibbonCallbackImpl;

typedef
IDispEventSimpleImpl<1, CConnect, &__uuidof(ApplicationEvents)>
ApplicationEventSink;

class ATL_NO_VTABLE CConnect
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CConnect, &__uuidof(Connect)>
	, public IDTExtensibilityImpl
	, public FormRegionStartupImpl
	, public ICustomTaskPaneConsumerImpl
	, public IRibbonExtensibilityImpl
	, public IRibbonCallbackImpl
	, public ApplicationEventSink
{
public:
	CConnect();

	// Override IDispatch Invoke
	STDMETHOD(Invoke)(DISPID dispidMember
		, const IID &riid
		, LCID lcid
		, WORD wFlags
		, DISPPARAMS *pdispparams
		, VARIANT *pvarResult
		, EXCEPINFO *pexceptinfo
		, UINT *puArgErr);

	// Setup the registration found in addin.rgs
	static HRESULT WINAPI UpdateRegistry(BOOL bRegister) throw()
	{
		ATL::_ATL_REGMAP_ENTRY regMapEntries[] =
		{ { OLESTR("PROGID"), ADDIN_PROGID }
		, { OLESTR("CLSID"), ADDIN_CLSID_STR }
		, { OLESTR("TYPELIB"), TYPELIB_GUID_STR }
		, { NULL, NULL }
		};

		return ATL::_pAtlModule->UpdateRegistryFromResource(IDR_ADDIN, bRegister, regMapEntries);
	}

	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(_FormRegionStartup)
		COM_INTERFACE_ENTRY(ICustomTaskPaneConsumer)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
		COM_INTERFACE_ENTRY(IRibbonCallback)
	END_COM_MAP()

	static _ATL_FUNC_INFO DispatchFuncInfo;

	BEGIN_SINK_MAP(CConnect)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents), dispidEventOptionsPagesAdd, OptionsPagesAdd, &DispatchFuncInfo)
	END_SINK_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct() { return S_OK; }
	void FinalRelease() { }

public:
	// IDTExtensibility2 interface
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom);

	// FormRegionStartup interface
	STDMETHOD(GetFormRegionStorage)(BSTR, IDispatch*, long, OlFormRegionMode, OlFormRegionSize, VARIANT*);
	STDMETHOD(BeforeFormRegionShow)(_FormRegion*);
	STDMETHOD(GetFormRegionManifest)(BSTR, long, VARIANT *);
	STDMETHOD(GetFormRegionIcon)(BSTR, long, OlFormRegionIcon, VARIANT*);

	// ICustomTaskPaneConsumer interface
	STDMETHOD(CTPFactoryAvailable)(ICTPFactory* CTPFactoryInst);

	// IRibbonExtensibility interface
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml);

	// IRibbonCallback Methods
	STDMETHOD(Button1Clicked)(IDispatch* ribbonControl);

	// ApplicationEvents Methods
	STDMETHOD(OptionsPagesAdd)(IDispatch* propertyPages);

private:
	STDMETHOD(HrCreateSampleTaskPane)(void);

	_ApplicationPtr m_pApplication;
	CComPtr<ICTPFactory> m_pCTPFactory;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
