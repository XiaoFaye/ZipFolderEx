/****************************** Module Header ******************************\
Module Name:  FileContextMenuExt.cpp
Project:      CppShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The code sample demonstrates creating a Shell context menu handler with C++. 

A context menu handler is a shell extension handler that adds commands to an 
existing context menu. Context menu handlers are associated with a particular 
file class and are called any time a context menu is displayed for a member 
of the class. While you can add items to a file class context menu with the 
registry, the items will be the same for all members of the class. By 
implementing and registering such a handler, you can dynamically add items to 
an object's context menu, customized for the particular object.

The example context menu handler adds the menu item "Display File Name (C++)"
to the context menu when you right-click a .cpp file in the Windows Explorer. 
Clicking the menu item brings up a message box that displays the full path 
of the .cpp file.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "ContextMenuExtractTo.h"
#include <string>
#include <strsafe.h>
#include <memory>
#include <Shlwapi.h>
#pragma comment(lib, "shlwapi.lib")
#include <comutil.h>
#pragma comment( lib, "comsuppwd.lib")

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

ContextMenuExtractTo::ContextMenuExtractTo(void) : m_cRef(1),
    m_pszMenuText(L"&Extract to"),
    m_pszVerb("cppdisplay"),
    m_pwszVerb(L"cppdisplay"),
    m_pszVerbCanonicalName("CppDisplayFileName"),
    m_pwszVerbCanonicalName(L"CppDisplayFileName"),
    m_pszVerbHelpText("Display File Name (C++)"),
    m_pwszVerbHelpText(L"Display File Name (C++)")
{
    InterlockedIncrement(&g_cDllRef);

	HRESULT hr;
	sfi = { 0 };
	hr = SHGetFileInfo(L".zip",
		FILE_ATTRIBUTE_NORMAL,
		&sfi,
		sizeof(sfi),
		SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);

	if (SUCCEEDED(hr))
	{
		hBitmap = BitmapFromIcon(sfi.hIcon);
	}

	ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&osvi);
}

ContextMenuExtractTo::~ContextMenuExtractTo(void)
{
	if (sfi.hIcon != NULL)
		DestroyIcon(sfi.hIcon);
	if (hBitmap != NULL)
		DeleteObject(hBitmap);
	
    InterlockedDecrement(&g_cDllRef);
}


void ContextMenuExtractTo::UnZipFile(BSTR strSrc, BSTR strDest)
{
	HRESULT hResult = S_FALSE;
	IShellDispatch *pIShellDispatch = NULL;
	Folder *pToFolder = NULL;
	VARIANT variantDir, variantFile, variantOpt;

	CoInitialize(NULL);

	try
	{
		hResult = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER,
			IID_IShellDispatch, (void **)&pIShellDispatch);
		if (SUCCEEDED(hResult))
		{
			VariantInit(&variantDir);
			variantDir.vt = VT_BSTR;
			variantDir.bstrVal = strDest;
			hResult = pIShellDispatch->NameSpace(variantDir, &pToFolder);

			if (SUCCEEDED(hResult))
			{
				Folder *pFromFolder = NULL;
				VariantInit(&variantFile);
				variantFile.vt = VT_BSTR;
				variantFile.bstrVal = strSrc;
				pIShellDispatch->NameSpace(variantFile, &pFromFolder);

				FolderItems *fi = NULL;
				pFromFolder->Items(&fi);

				VariantInit(&variantOpt);
				variantOpt.vt = VT_I4;
				variantOpt.lVal = 0;//FOF_NO_UI;

				VARIANT newV;
				VariantInit(&newV);
				newV.vt = VT_DISPATCH;
				newV.pdispVal = fi;
				hResult = pToFolder->CopyHere(newV, variantOpt);

				pFromFolder->Release();
				pToFolder->Release();
			}
			pIShellDispatch->Release();
		}
	}
	catch (DWORD code)
	{
		ShowMessage(code);
	}

	CoUninitialize();
}

void ContextMenuExtractTo::ShowMessage(DWORD code)
{
	TCHAR err[500];
	FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, code,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), err, 255, NULL);

	MessageBox(NULL, err, L"error", 0);
}

HBITMAP ContextMenuExtractTo::BitmapFromIcon(HICON hIcon)
{
	ICONINFO info = { 0 };
	if (hIcon == NULL || !GetIconInfo(hIcon, &info) || !info.fIcon)
	{
		return NULL;
	}

	INT nWidth = 0;
	INT nHeight = 0;

	if (info.hbmColor != NULL)
	{
		BITMAP bmp = { 0 };
		GetObject(info.hbmColor, sizeof(bmp), &bmp);

		nWidth = bmp.bmWidth;
		nHeight = bmp.bmHeight;
	}


	if (info.hbmColor != NULL)
	{
		DeleteObject(info.hbmColor);
		info.hbmColor = NULL;
	}

	if (info.hbmMask != NULL)
	{
		DeleteObject(info.hbmMask);
		info.hbmMask = NULL;
	}

	if (nWidth <= 0 || nHeight <= 0)
	{
		return NULL;
	}

	INT nPixelCount = nWidth * nHeight;

	HDC dc = ::GetDC(NULL);
	INT* pData = NULL;
	HDC dcMem = NULL;
	HBITMAP hBmpOld = NULL;
	bool* pOpaque = NULL;
	HBITMAP dib = NULL;
	BOOL bSuccess = FALSE;

	do
	{
		BITMAPINFOHEADER bi = { 0 };
		bi.biSize = sizeof(BITMAPINFOHEADER);
		bi.biWidth = nWidth;
		bi.biHeight = -nHeight;
		bi.biPlanes = 1;
		bi.biBitCount = 32;
		bi.biCompression = BI_RGB;
		dib = CreateDIBSection(dc, (BITMAPINFO*)&bi, DIB_RGB_COLORS, (VOID**)&pData, NULL, 0);
		if (dib == NULL) break;

		memset(pData, 0, nPixelCount * 4);

		dcMem = CreateCompatibleDC(dc);
		if (dcMem == NULL) break;

		hBmpOld = (HBITMAP)SelectObject(dcMem, dib);
		::DrawIconEx(dcMem, 0, 0, hIcon, nWidth, nHeight, 0, NULL, DI_MASK);

		pOpaque = new(std::nothrow) bool[nPixelCount];
		if (pOpaque == NULL) break;
		for (INT i = 0; i < nPixelCount; ++i)
		{
			pOpaque[i] = !pData[i];
		}

		memset(pData, 0, nPixelCount * 4);
		::DrawIconEx(dcMem, 0, 0, hIcon, nWidth, nHeight, 0, NULL, DI_NORMAL);

		BOOL bPixelHasAlpha = FALSE;
		UINT* pPixel = (UINT*)pData;
		for (INT i = 0; i<nPixelCount; ++i, ++pPixel)
		{
			if ((*pPixel & 0xff000000) != 0)
			{
				bPixelHasAlpha = TRUE;
				//MessageBox(L"has alpha");
				break;
			}
		}

		if (!bPixelHasAlpha)
		{
			pPixel = (UINT*)pData;
			for (INT i = 0; i <nPixelCount; ++i, ++pPixel)
			{
				if (pOpaque[i])
				{
					*pPixel |= 0xFF000000;
				}
				else
				{
					*pPixel &= 0x00FFFFFF;
				}
			}
		}

		bSuccess = TRUE;

	} while (FALSE);


	if (pOpaque != NULL)
	{
		delete[]pOpaque;
		pOpaque = NULL;
	}

	if (dcMem != NULL)
	{
		SelectObject(dcMem, hBmpOld);
		DeleteDC(dcMem);
	}

	::ReleaseDC(NULL, dc);

	if (!bSuccess)
	{
		if (dib != NULL)
		{
			DeleteObject(dib);
			dib = NULL;
		}
	}

	return dib;
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP ContextMenuExtractTo::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
		QITABENT(ContextMenuExtractTo, IContextMenu),
		QITABENT(ContextMenuExtractTo, IContextMenu2),
		QITABENT(ContextMenuExtractTo, IContextMenu3),
		QITABENT(ContextMenuExtractTo, IShellExtInit),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) ContextMenuExtractTo::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) ContextMenuExtractTo::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IShellExtInit

// Initialize the context menu handler.
IFACEMETHODIMP ContextMenuExtractTo::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    // The pDataObj pointer contains the objects being acted upon. In this 
    // example, we get an HDROP handle for enumerating the selected files and 
    // folders.
    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        // Get an HDROP handle.
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            // Determine how many files are involved in this operation. This 
            // code sample displays the custom context menu item when only 
            // one file is selected. 
            UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
            if (nFiles == 1)
            {
                // Get the path of the file.
                if (0 != DragQueryFile(hDrop, 0, m_szSelectedFile, 
                    ARRAYSIZE(m_szSelectedFile)))
                {
                    hr = S_OK;
                }
            }

            GlobalUnlock(stm.hGlobal);
        }

        ReleaseStgMedium(&stm);
    }

    // If any value other than S_OK is returned from the method, the context 
    // menu item is not displayed.
    return hr;
}

#pragma endregion


#pragma region IContextMenu

STDMETHODIMP ContextMenuExtractTo::HandleMenuMsg(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	// res is a dummy LRESULT variable.  It's not used (IContextMenu2::HandleMenuMsg()
	// doesn't provide a way to return values), but it's here so that the code 
	// in MenuMessageHandler() can be the same regardless of which interface it
	// gets called through (IContextMenu2 or 3).
	LRESULT res;

	return MenuMessageHandler(uMsg, wParam, lParam, &res);
}

STDMETHODIMP ContextMenuExtractTo::HandleMenuMsg2(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	// For messages that have no return value, pResult is NULL.
	// If it is NULL, I create a dummy LRESULT variable, so the code in
	// MenuMessageHandler() will always have a valid pResult pointer.

	if (NULL == pResult)
	{
		LRESULT res;
		return MenuMessageHandler(uMsg, wParam, lParam, &res);
	}
	else
		return MenuMessageHandler(uMsg, wParam, lParam, pResult);
}

HRESULT ContextMenuExtractTo::MenuMessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	switch (uMsg)
	{
	case WM_MEASUREITEM:
		return OnMeasureItem((MEASUREITEMSTRUCT*)lParam, pResult);
		break;

	case WM_DRAWITEM:
		return OnDrawItem((DRAWITEMSTRUCT*)lParam, pResult);
		break;
	}

	return S_OK;
}

HRESULT ContextMenuExtractTo::OnMeasureItem(MEASUREITEMSTRUCT* pmis, LRESULT* pResult)
{
	//Set the measures of the item.
	if ((pmis != NULL) && (pmis->CtlType == ODT_MENU)) {
		pmis->itemWidth += 2;
		pmis->itemHeight = 16;
	
		*pResult = true;
	}

	return S_OK;
}

HRESULT ContextMenuExtractTo::OnDrawItem(DRAWITEMSTRUCT* pdis, LRESULT* pResult)
{
	if ((pdis == NULL) || (pdis->CtlType != ODT_MENU))
		return S_OK; // not for a menu 
	
	if (sfi.hIcon == NULL) 
		return S_OK; 

	DrawIconEx(pdis->hDC, pdis->rcItem.left - 16, pdis->rcItem.top + (pdis->rcItem.bottom - pdis->rcItem.top - 16) / 2, sfi.hIcon, 16, 16, 0, NULL, DI_NORMAL);
	*pResult = TRUE;

	return S_OK;
}

//
//   FUNCTION: FileContextMenuExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP ContextMenuExtractTo::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    // Use either InsertMenu or InsertMenuItem to add menu items.
    // Learn how to add sub-menu from:
    // http://www.codeproject.com/KB/shell/ctxextsubmenu.aspx


	MENUITEMINFO mii = { sizeof(mii) };
	mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
	mii.wID = idCmdFirst + IDM_DISPLAY;
	if (osvi.dwMajorVersion < 6) mii.fType = MFT_OWNERDRAW;
	mii.dwTypeData = L"Extract Here";
	mii.fState = MFS_ENABLED;
	osvi.dwMajorVersion < 6 ? mii.hbmpItem = HBMMENU_CALLBACK : mii.hbmpItem = hBitmap;
	if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	MENUITEMINFO mii2 = { sizeof(mii2) };
	mii2.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
	mii2.wID = idCmdFirst + IDM_DISPLAY + 1;
	if (osvi.dwMajorVersion < 6) mii2.fType = MFT_OWNERDRAW;

	
	TCHAR extractTo[MAX_PATH] = L"Extract to \"";
	TCHAR selectedFile[MAX_PATH] = L"";
	StringCchCopy(selectedFile, MAX_PATH, this->m_szSelectedFile);
	PWSTR folder = PathFindFileName(selectedFile);
	PathRemoveExtension(folder);
	StringCchCat(extractTo, MAX_PATH, folder);
	StringCchCat(extractTo, MAX_PATH, L"\\\"");

	mii2.dwTypeData = extractTo;
	mii2.fState = MFS_ENABLED;
	osvi.dwMajorVersion < 6 ? mii2.hbmpItem = HBMMENU_CALLBACK : mii2.hbmpItem = hBitmap;
	if (!InsertMenuItem(hMenu, indexMenu+1, TRUE, &mii2))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 2, TRUE, &sep))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
    // Set the code value to the offset of the largest command identifier 
    // that was assigned, plus one (1).
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 2));
}


//
//   FUNCTION: FileContextMenuExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP ContextMenuExtractTo::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
    // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
    // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
    // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
    // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
    // and has additional members that allow Unicode strings to be passed.
    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    // Determines whether the command is identified by its offset or verb.
    // There are two ways to identify commands:
    // 
    //   1) The command's verb string 
    //   2) The command's identifier offset
    // 
    // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
    // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
    // holds a verb string. If the high-order word is zero, the command 
    // offset is in the low-order word of lpcmi->lpVerb.

    // For the ANSI case, if the high-order word is not zero, the command's 
    // verb string is in lpcmi->lpVerb. 
    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            //OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // For the Unicode case, if the high-order word is not zero, the 
    // command's verb string is in lpcmi->lpVerbW. 
    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            //OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // If the command cannot be identified through the verb string, then 
    // check the identifier offset.
    else
    {
		try
		{
			// Is the command identifier offset supported by this context menu 
			// extension?
			if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
			{
				TCHAR DestPath[MAX_PATH];
				StringCchCopy(DestPath, MAX_PATH, this->m_szSelectedFile);
				PathRemoveFileSpec(DestPath);
				UnZipFile(this->m_szSelectedFile, DestPath);
			}
			else if (LOWORD(pici->lpVerb) == IDM_DISPLAY + 1)
			{
				TCHAR DestPath[MAX_PATH];
				StringCchCopy(DestPath, MAX_PATH, this->m_szSelectedFile);
				PathRemoveExtension(DestPath);

				if (!PathFileExists(DestPath))
					CreateDirectory(DestPath, NULL);

				UnZipFile(this->m_szSelectedFile, DestPath);
			}
			else
			{
				// If the verb is not recognized by the context menu handler, it 
				// must return E_FAIL to allow it to be passed on to the other 
				// context menu handlers that might implement that verb.
				return E_FAIL;
			}
		}
		catch (DWORD code)
		{
			ShowMessage(code);
		}
    }

    return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP ContextMenuExtractTo::GetCommandString(UINT_PTR idCommand,
    UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

#pragma endregion
