
// TxtToXlsxDlg.cpp : 구현 파일
//

#include "stdafx.h"
#include "TxtToXlsx.h"
#include "TxtToXlsxDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CTxtToXlsxDlg 대화 상자



CTxtToXlsxDlg::CTxtToXlsxDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_TXTTOXLSX_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTxtToXlsxDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_PROGRESS_CONTINUE, m_ProgressCtrl);
}

BEGIN_MESSAGE_MAP(CTxtToXlsxDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_SELECT, &CTxtToXlsxDlg::OnBnClickedButtonSelect)
	ON_BN_CLICKED(IDC_BUTTON_CONVERT, &CTxtToXlsxDlg::OnBnClickedButtonConvert)
END_MESSAGE_MAP()


// CTxtToXlsxDlg 메시지 처리기

BOOL CTxtToXlsxDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	m_HDL = CreateEvent(NULL, TRUE, FALSE, NULL);

	// TODO: 여기에 추가 초기화 작업을 추가합니다.

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CTxtToXlsxDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CTxtToXlsxDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CTxtToXlsxDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

UINT CTxtToXlsxDlg::solve(LPVOID v)
{
	CTxtToXlsxDlg* p = (CTxtToXlsxDlg*)v;

	Aspose::Cells::Startup();

	string line;

	CString Input_File;
	CString Output_Folder;

	p->GetDlgItemTextW(IDC_EDIT_ORIGIN_FILE, Input_File);
	p->GetDlgItemTextW(IDC_EDIT_CONVERTED_FOLDER, Output_Folder);
	Output_Folder += "\\output.xlsx";

	U16String conv1_Output_Folder = string(CT2CA(Output_Folder)).c_str();

	ifstream file(Input_File); // 원하는 파일 경로 입력

	Workbook workbook;

	Worksheet worksheet = workbook.GetWorksheets().Get(0);

	char end_col = 'A';

	if (file.is_open())
	{
		int row = 1;

		while (getline(file, line))
		{
			string unit;

			stringstream sstream(line);

			char col = 'A';

			while (getline(sstream, unit, '\t'))
			{
				string A = col + to_string(row);

				U16String conv1 = A.c_str();
				U16String conv2 = unit.c_str();

				worksheet.GetCells().Get(conv1).PutValue(conv2);

				col++;
			}

			end_col = col - 1;

			row++;
		}
		file.close();
	}
	else
	{
		SetEvent(p->m_HDL);

		return -1;
	}

	worksheet.AutoFitColumns();

	string RangeTmp = "A1:";
	RangeTmp += end_col;
	RangeTmp += '1';

	worksheet.GetAutoFilter().SetRange(RangeTmp.c_str());

	workbook.Save(conv1_Output_Folder);

	Aspose::Cells::Cleanup();

	SetEvent(p->m_HDL);

	return 0;
}

UINT CTxtToXlsxDlg::waiting(LPVOID v)
{
	CTxtToXlsxDlg* p = (CTxtToXlsxDlg*)v;

	DWORD ret = WaitForSingleObject(p->m_HDL, INFINITE);

	p->m_ProgressCtrl.ModifyStyle(PBS_MARQUEE, 0);
	p->m_ProgressCtrl.SetPos(0);
	ResetEvent(p->m_HDL);

	if (ret == 0)
	{
		if (AfxMessageBox(_T("변환에 성공했습니다!")) == IDOK)
		{
			p->GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(1);
			p->GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(1);
		}
	}
	else if (ret == -1)
	{
		if (AfxMessageBox(_T("변환에 실패했습니다..")) == IDOK)
		{
			p->GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(1);
			p->GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(1);
		}
	}

	

	return 0;
}

void CTxtToXlsxDlg::OnBnClickedButtonSelect()
{
	CString str = _T("All files(*.*)|*.*|"); // 모든 파일 표시
											 // _T("Excel 파일 (*.xls, *.xlsx) |*.xls; *.xlsx|"); 와 같이 확장자를 제한하여 표시할 수 있음
	CFileDialog dlg(TRUE, _T("*.dat"), NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, str, this);

	if (dlg.DoModal() == IDOK)
	{
		CString strPathName = dlg.GetPathName();

		CString strStoredPath = strPathName.Left(strPathName.ReverseFind('\\'));

		// 파일 경로를 가져와 사용할 경우, Edit Control에 값 저장
		SetDlgItemText(IDC_EDIT_ORIGIN_FILE, strPathName);

		SetDlgItemText(IDC_EDIT_CONVERTED_FOLDER, strStoredPath);
	}
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}


void CTxtToXlsxDlg::OnBnClickedButtonConvert()
{
	m_ProgressCtrl.ModifyStyle(0, PBS_MARQUEE);
	m_ProgressCtrl.SetMarquee(1, 30);
	GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(0);
	GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(0);

	m_pWaitingThread = AfxBeginThread(waiting, this);
	m_pSolveThread = AfxBeginThread(solve, this);
	
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}