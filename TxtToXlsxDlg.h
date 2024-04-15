
// TxtToXlsxDlg.h : 헤더 파일
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


// CTxtToXlsxDlg 대화 상자
class CTxtToXlsxDlg : public CDialogEx
{
// 생성입니다.
public:
	CTxtToXlsxDlg(CWnd* pParent = NULL);	// 표준 생성자입니다.

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_TXTTOXLSX_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


// 구현입니다.
protected:
	HICON m_hIcon;

	// 생성된 메시지 맵 함수
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
