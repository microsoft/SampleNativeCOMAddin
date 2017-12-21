#include "ExplorerEventsSink.h"
#include "MAPI\TestMAPI.h"

_ATL_FUNC_INFO ExplorerEventsSink::FolderSwitchInfo = { CC_STDCALL, VT_EMPTY, 0, 0 };
_ATL_FUNC_INFO ExplorerEventsSink::OnCloseInfo = { CC_STDCALL, VT_EMPTY, 0, 0 };

ExplorerEventsSink::ExplorerEventsSink(_ExplorerPtr piExplorer)
{
	m_piExplorer = piExplorer;
	DispEventAdvise(m_piExplorer, &__uuidof(Outlook::ExplorerEvents));
}

ExplorerEventsSink::~ExplorerEventsSink()
{
	DispEventUnadvise((IUnknown*)m_piExplorer);
}

void ExplorerEventsSink::OnClose()
{
	//MessageBoxW(NULL, L"OnClose", L"Sample Add-In", MB_OK | MB_ICONINFORMATION);
}

void ExplorerEventsSink::FolderSwitch()
{
	//MessageBoxW(NULL, L"FolderSwitch", L"Sample Add-In", MB_OK | MB_ICONINFORMATION);
	//TestMAPI::TestInbox(L"FolderSwitch", true);
}