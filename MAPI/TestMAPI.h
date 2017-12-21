#pragma once
#include "..\stdafx.h"
#include "MAPIx.h"
#include <string>

namespace TestMAPI
{
	// IMAPISecureMessage
	DEFINE_GUID(IID_IMAPISecureMessage, 0x253cc320, 0xeab6, 0x11d0, 0x82, 0x22, 0, 0x60, 0x97, 0x93, 0x87, 0xea);

	struct IMAPISecureMessage : IUnknown
	{
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder1() = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder2() = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder3() = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder4() = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder5() = 0;
		virtual HRESULT STDMETHODCALLTYPE GetBaseMessage(LPMESSAGE FAR * ppmsg) = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder7() = 0;
		virtual HRESULT STDMETHODCALLTYPE PlaceHolder8() = 0;
	};

	void TestInbox(std::wstring caller, bool bOnline);
	void TestGetProp(LPMAPIPROP lpObj);
	void TestBaseMessage(IMAPISecureMessage* lpObj);
	void TestSecureMessage(LPMAPIPROP lpObj);
};