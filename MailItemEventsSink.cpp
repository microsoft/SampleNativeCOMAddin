#include "MailItemEventsSink.h"
#include "MAPI/TestMAPI.h"

_ATL_FUNC_INFO MailItemEventsSink::UnloadInfo = { CC_STDCALL, VT_EMPTY, 0, 0 };

MailItemEventsSink::MailItemEventsSink(_MailItemPtr piMailItem)
{
	m_piMailItem = piMailItem;
	DispEventAdvise(m_piMailItem, &__uuidof(ItemEvents_10));
}

MailItemEventsSink::~MailItemEventsSink()
{
	DispEventUnadvise(static_cast<IUnknown*>(m_piMailItem));
}

void MailItemEventsSink::Unload()
{
	//delete this;
}