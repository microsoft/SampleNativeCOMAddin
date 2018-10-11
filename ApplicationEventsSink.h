#pragma once
#include "stdafx.h"

class ApplicationEventsSink;

typedef IDispEventSimpleImpl<1, ApplicationEventsSink, &__uuidof(ApplicationEvents_11)> IApplicationEventsSink;

class ApplicationEventsSink :
	public IApplicationEventsSink
{
public:
	ApplicationEventsSink(_ApplicationPtr piApp);
	virtual ~ApplicationEventsSink();

	static _ATL_FUNC_INFO OptionsPagesAddInfo;
	static _ATL_FUNC_INFO MapiLogonCompleteInfo;
	static _ATL_FUNC_INFO ItemSendInfo;
	static _ATL_FUNC_INFO ItemLoadInfo;
	static _ATL_FUNC_INFO ItemUnloadInfo;

	BEGIN_SINK_MAP(ApplicationEventsSink)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), dispidEventOptionsPagesAdd, OptionsPagesAdd, &OptionsPagesAddInfo)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), dispidEventMapiLogonComplete, MapiLogonComplete, &MapiLogonCompleteInfo)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), dispidEventItemSend, ItemSend, &ItemSendInfo)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), dispidEventItemLoad, ItemLoad, &ItemLoadInfo)
	END_SINK_MAP()

	// ApplicationEvents Methods
	STDMETHOD(OptionsPagesAdd)(IDispatch* pages);
	STDMETHOD(MapiLogonComplete)();
	STDMETHOD(ItemSend)(IDispatch* Item, VARIANT_BOOL* Cancel);
	STDMETHOD(ItemLoad)(IDispatch* Item);

private:
	_ApplicationPtr m_piApp = nullptr;
};

