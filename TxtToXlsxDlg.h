
// TxtToXlsxDlg.h : ��� ����
//

#pragma once

#include <iostream>
#include <fstream>
#include <vector>
#include <stack>
#include <sstream>
#include <io.h>

#include <Windows.h>
#include "Aspose.Cells.h"
#include "afxcmn.h"

using namespace Aspose::Cells;
using namespace std;


// CTxtToXlsxDlg ��ȭ ����
class CTxtToXlsxDlg : public CDialogEx
{
// �����Դϴ�.
public:
	CTxtToXlsxDlg(CWnd* pParent = NULL);	// ǥ�� �������Դϴ�.

// ��ȭ ���� �������Դϴ�.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_TXTTOXLSX_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �����Դϴ�.


// �����Դϴ�.
protected:
	HICON m_hIcon;

	// ������ �޽��� �� �Լ�
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonSelect();
	afx_msg void OnBnClickedButtonConvert();

	static UINT solve(LPVOID);
	static UINT waiting(LPVOID);

	CWinThread* m_pSolveThread;
	CWinThread* m_pWaitingThread;
	HANDLE m_HDL;

	CProgressCtrl m_ProgressCtrl;
};
