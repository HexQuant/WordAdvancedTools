#include "pch.h"
#include "RevisionSettingsDialog.h"


// CRevisionSettingDialog


LRESULT CRevisionSettingsDialog::OnBnClickedButton1(WORD /*wNotifyCode*/, WORD wID, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CWindow pEdit = GetDlgItem(IDC_EDIT2);
	pEdit.GetWindowText(rm->Text);
	GetDlgItem(IDC_COMBO1).GetWindowText(rm->StyleName);

	//auto i = GetDlgItem(IDC_CHECK1).GetCheck();
	if (IsDlgButtonChecked(IDC_CHECK1) == BST_CHECKED)
	{
		rm->IsField = true;
	}
	else
	{
		rm->IsField = false;

	}
	auto str = SysAllocString(L"0");

	GetDlgItem(IDC_EDIT3).GetWindowText(str);
	rm->PortraitFirstMargin = _wtof(str);

	GetDlgItem(IDC_EDIT4).GetWindowText(str);
	rm->LandscapeFirstMargin = _wtof(str);

	GetDlgItem(IDC_EDIT5).GetWindowText(str);
	rm->PortraitSecondMargin = _wtof(str);

	GetDlgItem(IDC_EDIT6).GetWindowText(str);
	rm->LandscapeSecondMargin = _wtof(str);

	EndDialog(wID);

	return 0;
}


LRESULT CRevisionSettingsDialog::OnEnUpdateEdit2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the __super::OnInitDialog()
	// function to send the EM_SETEVENTMASK message to the control
	// with the ENM_UPDATE flag ORed into the lParam mask.

	// TODO:  Add your control notification handler code here
	return 0;
}


LRESULT CRevisionSettingsDialog::OnEnChangeEdit2(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the __super::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here

	return 0;
}
