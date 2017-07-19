#include <tchar.h>
#include <mapix.h>
#include <mapiutil.h>
#include <mapitags.h>
#include "defaults.h"

STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, ULONG ulExplicitFlags, LPMDB * lpMDB)
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpStoresTbl = NULL;

	enum { EID, NAME, NUM_COLS };
	static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS, PR_ENTRYID, PR_DISPLAY_NAME };

	if (!lpMAPISession || !lpMDB)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	*lpMDB = NULL;

	//Get the table of all the message stores available
	hRes = lpMAPISession->GetMsgStoresTable(0, &lpStoresTbl);
	if (SUCCEEDED(hRes) && lpStoresTbl)
	{
		SRestriction	sres;
		SPropValue		spv;
		LPSRowSet		pRow = NULL;

		//Set up restriction for the default store
		sres.rt = RES_PROPERTY;								//Comparing a property
		sres.res.resProperty.relop = RELOP_EQ;				//Testing equality
		sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE;	//Tag to compare
		sres.res.resProperty.lpProp = &spv;					//Prop tag and value to compare against

		spv.ulPropTag = PR_DEFAULT_STORE;	//Tag type
		spv.Value.b = TRUE;				//Tag value

										//Convert the table to an array which can be stepped through
										//Only one message store should have PR_DEFAULT_STORE set to true, so only one will be returned
		hRes = HrQueryAllRows(
			lpStoresTbl,					//Table to query
			(LPSPropTagArray)&sptCols,	//Which columns to get
			&sres,						//Restriction to use
			NULL,						//No sort order
			0,							//Max number of rows (0 means no limit)
			&pRow);						//Array to return
		if (SUCCEEDED(hRes) && pRow && pRow->cRows == 1)
		{
			LPMDB	lpTempMDB = NULL;
			//Open the first returned (default) message store
			hRes = lpMAPISession->OpenMsgStore(
				NULL,//Window handle for dialogs
				pRow->aRow[0].lpProps[EID].Value.bin.cb,//size and...
				(LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb,//value of entry to open
				NULL,//Use default interface (IMsgStore) to open store
				MAPI_BEST_ACCESS | ulExplicitFlags,//Flags
				&lpTempMDB);//Pointer to place the store in
			if (SUCCEEDED(hRes) && lpTempMDB)
			{
				//Assign the out parameter
				*lpMDB = lpTempMDB;
			}
		}

		FreeProws(pRow);
	}

	if (lpStoresTbl)
		lpStoresTbl->Release();

	return hRes;
}

STDMETHODIMP OpenDefaultByProp(ULONG ulPropTag, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpRoot = NULL;
	LPMAPIFOLDER	lpInbox = NULL;
	ULONG			ulObjType = 0;
	LPSPropValue	lpEIDProp = NULL;

	*lpFolder = NULL;

	hRes = lpMDB->OpenEntry(0, NULL, NULL, MAPI_BEST_ACCESS,
		&ulObjType, (LPUNKNOWN*)&lpRoot);

	if (SUCCEEDED(hRes) && lpRoot)
	{
		// Get the entry id from the root folder
		hRes = HrGetOneProp(lpRoot, ulPropTag, &lpEIDProp);
		if (MAPI_E_NOT_FOUND == hRes)
		{
			// Ok..just to confuse things, Outlook 11 moved the props
			// to the Inbox folder. Technically that's legal, so we 
			// have to handle it. If this isn't found, try the inbox.
			hRes = OpenInbox(lpMDB, NULL, &lpInbox);
			if (SUCCEEDED(hRes) && lpInbox)
			{
				hRes = HrGetOneProp(lpInbox, ulPropTag, &lpEIDProp);
			}
		}
	}

	// Open whatever folder we got..
	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		LPMAPIFOLDER	lpTemp = NULL;

		hRes = lpMDB->OpenEntry(
			lpEIDProp->Value.bin.cb,
			(LPENTRYID)lpEIDProp->Value.bin.lpb,
			NULL,
			ulExplicitFlags,
			&ulObjType,
			(LPUNKNOWN*)&lpTemp);
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	if (lpInbox) lpInbox->Release();
	if (lpRoot) lpRoot->Release();
	return hRes;
}

STDMETHODIMP OpenPropFromMDB(ULONG ulPropTag, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	LPSPropValue	lpEIDProp = NULL;

	*lpFolder = NULL;

	hRes = HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp);

	// Open whatever folder we got..
	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		LPMAPIFOLDER	lpTemp = NULL;

		hRes = lpMDB->OpenEntry(
			lpEIDProp->Value.bin.cb,
			(LPENTRYID)lpEIDProp->Value.bin.lpb,
			NULL,
			ulExplicitFlags,
			&ulObjType,
			(LPUNKNOWN*)&lpTemp);
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	return hRes;
}

STDMETHODIMP OpenDefaultFolder(ULONG ulFolder, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;

	if (!lpMDB || !lpFolder)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	*lpFolder = NULL;
	if (!ulExplicitFlags) ulExplicitFlags = MAPI_BEST_ACCESS;

	switch (ulFolder)
	{
	case DEFAULT_CALENDAR:
		hRes = OpenDefaultByProp(PR_CALENDAR_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_CONTACTS:
		hRes = OpenDefaultByProp(PR_CONTACTS_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_JOURNAL:
		hRes = OpenDefaultByProp(PR_JOURNAL_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_NOTES:
		hRes = OpenDefaultByProp(PR_NOTES_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_TASKS:
		hRes = OpenDefaultByProp(PR_TASKS_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_REMINDERS:
		hRes = OpenDefaultByProp(PR_REMINDERS_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_DRAFTS:
		hRes = OpenDefaultByProp(PR_DRAFTS_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_SENTITEMS:
		hRes = OpenPropFromMDB(PR_IPM_SENTMAIL_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_OUTBOX:
		hRes = OpenPropFromMDB(PR_IPM_OUTBOX_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_DELETEDITEMS:
		hRes = OpenPropFromMDB(PR_IPM_WASTEBASKET_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_FINDER:
		hRes = OpenPropFromMDB(PR_FINDER_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_IPM_SUBTREE:
		hRes = OpenPropFromMDB(PR_IPM_SUBTREE_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_INBOX:
		hRes = OpenInbox(lpMDB, ulExplicitFlags, lpFolder);
		break;
	case DEFAULT_COMMON_VIEWS:
		hRes = OpenPropFromMDB(PR_COMMON_VIEWS_ENTRYID, lpMDB, ulExplicitFlags, lpFolder);
		break;
	default:
		hRes = MAPI_E_INVALID_PARAMETER;
	}

	return hRes;
}

STDMETHODIMP OpenInbox(LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpInboxFolder)
{
	HRESULT			hRes = S_OK;
	ULONG			cbInbox = 0;
	LPENTRYID		lpbInbox = NULL;

	if (!lpMDB || !lpInboxFolder)
		return MAPI_E_INVALID_PARAMETER;

	*lpInboxFolder = NULL;

	hRes = lpMDB->GetReceiveFolder(
		_T("IPM.Note"),	//Get default receive folder
		fMapiUnicode,	//Flags
		&cbInbox,		//Size and ...
		&lpbInbox,		//Value of the EntryID to be returned
		NULL);			//We don't care to see the class returned
	if (SUCCEEDED(hRes) && cbInbox && lpbInbox)
	{
		ULONG			ulObjType = 0;
		LPMAPIFOLDER	lpTempFolder = NULL;

		hRes = lpMDB->OpenEntry(
			cbInbox,						//Size and...
			lpbInbox,						//Value of the Inbox's EntryID
			NULL,							//We want the default interface (IMAPIFolder)
			ulExplicitFlags,				// Flags
			&ulObjType,						//Object returned type
			(LPUNKNOWN *)&lpTempFolder);	//Returned folder
		if (SUCCEEDED(hRes) && lpTempFolder)
		{
			//Assign the out parameter
			*lpInboxFolder = lpTempFolder;
		}
	}

	//Always clean up your memory here!
	MAPIFreeBuffer(lpbInbox);
	return hRes;
}

STDMETHODIMP GetFirstMessage(LPMAPIFOLDER lpFolder, ULONG ulExplicitFlags, LPMESSAGE * lppMessage)
{
	return GetMessageNum(lpFolder, ulExplicitFlags, 0, lppMessage);
}

STDMETHODIMP GetMessageNum(LPMAPIFOLDER lpFolder, ULONG ulExplicitFlags, ULONG ulMsgNum, LPMESSAGE * lppMessage)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpContents = NULL;

	enum { EID, NUM_COLS };
	static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS, PR_ENTRYID };

	if (!lpFolder || !lppMessage)
		return MAPI_E_INVALID_PARAMETER;

	*lppMessage = NULL;

	hRes = lpFolder->GetContentsTable(0, &lpContents);
	if (SUCCEEDED(hRes) && lpContents)
	{
		hRes = lpContents->SetColumns((LPSPropTagArray)&sptCols, 0);
		if (SUCCEEDED(hRes))
		{
			hRes = lpContents->SeekRow(BOOKMARK_BEGINNING, ulMsgNum, NULL);
			if (SUCCEEDED(hRes))
			{
				LPSRowSet pRow = NULL;

				hRes = lpContents->QueryRows(1, 0, &pRow);
				if (SUCCEEDED(hRes) && pRow)
				{
					if (pRow->cRows < 1)
					{
						hRes = MAPI_E_TABLE_EMPTY;
					}
					else
					{
						if (pRow->aRow[0].lpProps[EID].ulPropTag == PR_ENTRYID)
						{
							LPMESSAGE lpTempMessage = NULL;
							ULONG ulObjType = 0;

							hRes = lpFolder->OpenEntry(pRow->aRow[0].lpProps[EID].Value.bin.cb,
								(LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb,
								NULL,
								ulExplicitFlags,
								&ulObjType,
								(LPUNKNOWN*)&lpTempMessage);
							if (SUCCEEDED(hRes) && lpTempMessage)
							{
								*lppMessage = lpTempMessage;
							}
						}
					}
				}

				FreeProws(pRow);
			}
		}
	}

	if (lpContents)lpContents->Release();

	return hRes;
}

HRESULT OpenMessageStoreGUID(
	LPMAPISESSION	lpMAPISession,
	LPCSTR lpGUID,
	LPMDB* lppMDB)
{
	LPMAPITABLE	pStoresTbl = NULL;
	LPSRowSet	pRow = NULL;
	ULONG		ulRowNum;
	HRESULT		hRes = S_OK;

	enum { EID, STORETYPE, NUM_COLS };
	static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS,
		PR_ENTRYID,
		PR_MDB_PROVIDER
	};

	*lppMDB = NULL;
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	hRes = lpMAPISession->GetMsgStoresTable(0, &pStoresTbl);

	if (SUCCEEDED(hRes) && pStoresTbl)
	{
		hRes = HrQueryAllRows(
			pStoresTbl,					//table to query
			(LPSPropTagArray)&sptCols,	//columns to get
			NULL,						//restriction to use
			NULL,						//sort order
			0,							//max number of rows
			&pRow);
		if (SUCCEEDED(hRes) && pRow)
		{
			for (ulRowNum = 0; ulRowNum < pRow->cRows; ulRowNum++)
			{
				hRes = S_OK;
				//check to see if we have a folder with a matching GUID
				if (IsEqualMAPIUID(
					pRow->aRow[ulRowNum].lpProps[STORETYPE].Value.bin.lpb,
					lpGUID))
				{
					hRes = lpMAPISession->OpenMsgStore(
						NULL,
						pRow->aRow[ulRowNum].lpProps[EID].Value.bin.cb,
						(LPENTRYID)pRow->aRow[ulRowNum].lpProps[EID].Value.bin.lpb,
						NULL,
						MDB_WRITE,
						lppMDB);
					break;
				}
			}
		}
	}
	if (!*lppMDB) hRes = MAPI_E_NOT_FOUND;

	if (pRow) FreeProws(pRow);
	if (pStoresTbl) pStoresTbl->Release();
	return hRes;
}

// Avoid slurping in edkmdb.h
#define pidStoreMin						0x6618
#define PR_IPM_PUBLIC_FOLDERS_ENTRYID			PROP_TAG( PT_BINARY, pidStoreMin+0x19)

HRESULT OpenPF(LPMDB lpPFMDB, LPCTSTR lpszPFName, LPMAPIFOLDER *lppPF)
{
	if (!lpPFMDB || !lpszPFName || !lppPF) return MAPI_E_INVALID_PARAMETER;
	*lppPF = NULL;

	HRESULT hRes = S_OK;
	LPSPropValue pPropAllRoot = NULL;
	hRes = HrGetOneProp(lpPFMDB, PR_IPM_PUBLIC_FOLDERS_ENTRYID, &pPropAllRoot);
	if (SUCCEEDED(hRes) && pPropAllRoot && PR_IPM_PUBLIC_FOLDERS_ENTRYID == pPropAllRoot->ulPropTag)
	{
		ULONG ulType = NULL;
		LPMAPIFOLDER lpRootFolder = NULL;
		hRes = lpPFMDB->OpenEntry(
			pPropAllRoot->Value.bin.cb,
			(LPENTRYID)pPropAllRoot->Value.bin.lpb,
			NULL,
			MAPI_MODIFY,
			&ulType,
			(LPUNKNOWN*)&lpRootFolder);
		if (lpRootFolder)
		{
			LPMAPITABLE lpFolderTable = NULL;
			//Get the table of all the folders
			hRes = lpRootFolder->GetHierarchyTable(0, &lpFolderTable);
			if (SUCCEEDED(hRes) && lpFolderTable)
			{
				static SizedSPropTagArray(2, sptCols) = { 2, PR_ENTRYID, PR_DISPLAY_NAME };
				SRestriction	sres = { 0 };
				SPropValue		spv = { 0 };
				LPSRowSet		pRow = NULL;

				//Set up restriction for the default store
				sres.rt = RES_PROPERTY;								//Comparing a property
				sres.res.resProperty.relop = RELOP_EQ;				//Testing equality
				sres.res.resProperty.ulPropTag = PR_DISPLAY_NAME;	//Tag to compare
				sres.res.resProperty.lpProp = &spv;					//Prop tag and value to compare against

				spv.ulPropTag = PR_DISPLAY_NAME;	//Tag type
				spv.Value.LPSZ = (LPTSTR)lpszPFName;		//Tag value

				hRes = HrQueryAllRows(
					lpFolderTable,					//Table to query
					(LPSPropTagArray)&sptCols,	//Which columns to get
					&sres,						//Restriction to use
					NULL,						//No sort order
					0,							//Max number of rows (0 means no limit)
					&pRow);						//Array to return
				if (SUCCEEDED(hRes) &&
					pRow &&
					pRow->cRows == 1 &&
					PR_ENTRYID == pRow->aRow[0].lpProps[0].ulPropTag)
				{
					LPMAPIFOLDER	lpRetFolder = NULL;
					hRes = lpPFMDB->OpenEntry(
						pRow->aRow[0].lpProps[0].Value.bin.cb,
						(LPENTRYID)pRow->aRow[0].lpProps[0].Value.bin.lpb,
						NULL,
						MAPI_BEST_ACCESS,
						&ulType,
						(LPUNKNOWN*)&lpRetFolder);
					if (SUCCEEDED(hRes) && lpRetFolder)
					{
						*lppPF = lpRetFolder;
					}
					else if (lpRetFolder) lpRetFolder->Release();
				}
				MAPIFreeBuffer(pRow);
			}
			if (lpFolderTable) lpFolderTable->Release();

			lpRootFolder->Release();
		}
	}
	return hRes;
}

HRESULT OpenFolder(LPMDB lpMDB, LPCTSTR lpszFolderName, LPMAPIFOLDER *lppFolder)
{
	if (!lpMDB || !lpszFolderName || !lppFolder) return MAPI_E_INVALID_PARAMETER;
	*lppFolder = NULL;

	HRESULT hRes = S_OK;

	ULONG ulType = NULL;
	LPMAPIFOLDER lpRootFolder = NULL;
	hRes = lpMDB->OpenEntry(
		0,
		0,
		NULL,
		MAPI_MODIFY,
		&ulType,
		(LPUNKNOWN*)&lpRootFolder);
	if (lpRootFolder)
	{
		LPMAPITABLE lpFolderTable = NULL;
		//Get the table of all the folders
		hRes = lpRootFolder->GetHierarchyTable(0, &lpFolderTable);
		if (SUCCEEDED(hRes) && lpFolderTable)
		{
			static SizedSPropTagArray(2, sptCols) = { 2, PR_ENTRYID, PR_DISPLAY_NAME };
			SRestriction	sres = { 0 };
			SPropValue		spv = { 0 };
			LPSRowSet		pRow = NULL;

			//Set up restriction for the default store
			sres.rt = RES_PROPERTY;								//Comparing a property
			sres.res.resProperty.relop = RELOP_EQ;				//Testing equality
			sres.res.resProperty.ulPropTag = PR_DISPLAY_NAME;	//Tag to compare
			sres.res.resProperty.lpProp = &spv;					//Prop tag and value to compare against

			spv.ulPropTag = PR_DISPLAY_NAME;	//Tag type
			spv.Value.LPSZ = (LPTSTR)lpszFolderName;		//Tag value

			hRes = HrQueryAllRows(
				lpFolderTable,					//Table to query
				(LPSPropTagArray)&sptCols,	//Which columns to get
				&sres,						//Restriction to use
				NULL,						//No sort order
				0,							//Max number of rows (0 means no limit)
				&pRow);						//Array to return
			if (SUCCEEDED(hRes) &&
				pRow &&
				pRow->cRows == 1 &&
				PR_ENTRYID == pRow->aRow[0].lpProps[0].ulPropTag)
			{
				LPMAPIFOLDER	lpRetFolder = NULL;
				hRes = lpMDB->OpenEntry(
					pRow->aRow[0].lpProps[0].Value.bin.cb,
					(LPENTRYID)pRow->aRow[0].lpProps[0].Value.bin.lpb,
					NULL,
					MAPI_BEST_ACCESS,
					&ulType,
					(LPUNKNOWN*)&lpRetFolder);
				if (SUCCEEDED(hRes) && lpRetFolder)
				{
					*lppFolder = lpRetFolder;
				}
				else if (lpRetFolder) lpRetFolder->Release();
			}
			MAPIFreeBuffer(pRow);
		}
		if (lpFolderTable) lpFolderTable->Release();

		lpRootFolder->Release();
	}

	return hRes;
}