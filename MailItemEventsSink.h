#pragma once
#include "stdafx.h"

class MailItemEventsSink;

class MailItemEventsSink :
	public IDispEventSimpleImpl<1, MailItemEventsSink, &__uuidof(ItemEvents_10)>
{
public:
	MailItemEventsSink(_MailItemPtr piMailItem);
	virtual ~MailItemEventsSink();

	static _ATL_FUNC_INFO UnloadInfo;

	BEGIN_SINK_MAP(MailItemEventsSink)
		SINK_ENTRY_INFO(1, __uuidof(ItemEvents_10), dispidEventItemUnload, Unload, &UnloadInfo)
	END_SINK_MAP()

	// MailItemEvents Methods
	void __stdcall Unload();

private:
	_MailItemPtr m_piMailItem = nullptr;
};