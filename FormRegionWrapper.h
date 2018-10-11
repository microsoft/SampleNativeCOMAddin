/*!-----------------------------------------------------------------------
	formregionwrapper.h
-----------------------------------------------------------------------!*/
#pragma once
#include "stdafx.h"

/*!-----------------------------------------------------------------------
	FormRegionWrapper- Used to help track all the form regions and to
	listen to some basic events such as the send button click and close.
-----------------------------------------------------------------------!*/
class FormRegionWrapper;

typedef
	IDispEventSimpleImpl<1, FormRegionWrapper, &__uuidof(OlkCommandButtonEvents)>
	OlkCommandButtonEventSink;

typedef
	IDispEventSimpleImpl<2, FormRegionWrapper, &__uuidof(FormRegionEvents)>
	FormRegionEventSink;

class FormRegionWrapper
	: public FormRegionEventSink,
 public OlkCommandButtonEventSink
{
public:
	static HRESULT Setup(_FormRegion* pFormRegion);

private:
	HRESULT HrInit(_FormRegion* pFormRegion);

	static _ATL_FUNC_INFO VoidFuncInfo;
public:
	BEGIN_SINK_MAP(FormRegionWrapper)
		SINK_ENTRY_INFO(1, __uuidof(OlkCommandButtonEvents), DISPID_CLICK, OnButton1Click, &VoidFuncInfo)
		SINK_ENTRY_INFO(2, __uuidof(FormRegionEvents), dispidEventOnClose, OnFormRegionClose, &VoidFuncInfo)
	END_SINK_MAP()

	void __stdcall OnButton1Click() const;
	void __stdcall OnFormRegionClose();

private:
	_FormRegionPtr m_spFormRegion = nullptr;
	_OlkCommandButtonPtr m_spOlkCmdBtn = nullptr;
	_MailItemPtr m_spMailItem = nullptr;
};

