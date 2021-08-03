#pragma once

#include "resource.h"
#include <atlhost.h>

#include "RevisionMgr.h"

using namespace ATL;

// CRevisionSettingsDialog

class CRevisionSettingsDialog :
	public CAxDialogImpl<CRevisionSettingsDialog>
{
public:

	RevisionMgr* rm;

	CRevisionSettingsDialog()
	{
	}

	~CRevisionSettingsDialog()
	{
	}

	enum { IDD = 107 };

BEGIN_MSG_MAP(CRevisionSettingsDialog)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	COMMAND_HANDLER(IDOK, BN_CLICKED, OnClickedOK)
	COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnClickedCancel)
	CHAIN_MSG_MAP(CAxDialogImpl<CRevisionSettingsDialog>)
	COMMAND_HANDLER(IDC_BUTTON1, BN_CLICKED, OnBnClickedButton1)
	COMMAND_HANDLER(IDC_EDIT2, EN_UPDATE, OnEnUpdateEdit2)
	COMMAND_HANDLER(IDC_EDIT2, EN_CHANGE, OnEnChangeEdit2)
END_MSG_MAP()

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		CAxDialogImpl<CRevisionSettingsDialog>::OnInitDialog(uMsg, wParam, lParam, bHandled);
		bHandled = TRUE;
		
		GetDlgItem(IDC_EDIT2).SetWindowText(rm->Text);
		GetDlgItem(IDC_COMBO1).SetWindowText(rm->StyleName);
	
		CheckDlgButton(IDC_CHECK1,rm->IsField);

		auto pf = std::to_wstring(rm->PortraitFirstMargin);
		GetDlgItem(IDC_EDIT3).SetWindowText(SysAllocStringLen(pf.c_str(), pf.length()));

		auto lf = std::to_wstring(rm->LandscapeFirstMargin);
		GetDlgItem(IDC_EDIT4).SetWindowText(SysAllocStringLen(lf.c_str(), lf.length()));

		auto ps = std::to_wstring(rm->PortraitSecondMargin);
		GetDlgItem(IDC_EDIT5).SetWindowText(SysAllocStringLen(ps.c_str(), ps.length()));

		auto ls = std::to_wstring(rm->LandscapeSecondMargin);
		GetDlgItem(IDC_EDIT6).SetWindowText(SysAllocStringLen(ls.c_str(), ls.length()));

		return 1;  // Let the system set the focus
	}

	LRESULT OnClickedOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		EndDialog(wID);
		return 0;
	}

	LRESULT OnClickedCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		EndDialog(wID);
		return 0;
	}
	LRESULT OnBnClickedButton1(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEnUpdateEdit2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEnChangeEdit2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
};
