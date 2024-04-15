
// TxtToXlsxDlg.cpp : ���� ����
//

#include "stdafx.h"
#include "TxtToXlsx.h"
#include "TxtToXlsxDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ���� ���α׷� ������ ���Ǵ� CAboutDlg ��ȭ �����Դϴ�.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// ��ȭ ���� �������Դϴ�.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV �����Դϴ�.

// �����Դϴ�.
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


// CTxtToXlsxDlg ��ȭ ����



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


// CTxtToXlsxDlg �޽��� ó����

BOOL CTxtToXlsxDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// �ý��� �޴��� "����..." �޴� �׸��� �߰��մϴ�.

	// IDM_ABOUTBOX�� �ý��� ��� ������ �־�� �մϴ�.
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

	// �� ��ȭ ������ �������� �����մϴ�.  ���� ���α׷��� �� â�� ��ȭ ���ڰ� �ƴ� ��쿡��
	//  �����ӿ�ũ�� �� �۾��� �ڵ����� �����մϴ�.
	SetIcon(m_hIcon, TRUE);			// ū �������� �����մϴ�.
	SetIcon(m_hIcon, FALSE);		// ���� �������� �����մϴ�.

	m_HDL = CreateEvent(NULL, TRUE, FALSE, NULL);

	// TODO: ���⿡ �߰� �ʱ�ȭ �۾��� �߰��մϴ�.

	return TRUE;  // ��Ŀ���� ��Ʈ�ѿ� �������� ������ TRUE�� ��ȯ�մϴ�.
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

// ��ȭ ���ڿ� �ּ�ȭ ���߸� �߰��� ��� �������� �׸�����
//  �Ʒ� �ڵ尡 �ʿ��մϴ�.  ����/�� ���� ����ϴ� MFC ���� ���α׷��� ��쿡��
//  �����ӿ�ũ���� �� �۾��� �ڵ����� �����մϴ�.

void CTxtToXlsxDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // �׸��⸦ ���� ����̽� ���ؽ�Ʈ�Դϴ�.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Ŭ���̾�Ʈ �簢������ �������� ����� ����ϴ�.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// �������� �׸��ϴ�.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// ����ڰ� �ּ�ȭ�� â�� ���� ���ȿ� Ŀ���� ǥ�õǵ��� �ý��ۿ���
//  �� �Լ��� ȣ���մϴ�.
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

	ifstream file(Input_File); // ���ϴ� ���� ��� �Է�

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
		if (AfxMessageBox(_T("��ȯ�� �����߽��ϴ�!")) == IDOK)
		{
			p->GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(1);
			p->GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(1);
		}
	}
	else if (ret == -1)
	{
		if (AfxMessageBox(_T("��ȯ�� �����߽��ϴ�..")) == IDOK)
		{
			p->GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(1);
			p->GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(1);
		}
	}

	

	return 0;
}

void CTxtToXlsxDlg::OnBnClickedButtonSelect()
{
	CString str = _T("All files(*.*)|*.*|"); // ��� ���� ǥ��
											 // _T("Excel ���� (*.xls, *.xlsx) |*.xls; *.xlsx|"); �� ���� Ȯ���ڸ� �����Ͽ� ǥ���� �� ����
	CFileDialog dlg(TRUE, _T("*.dat"), NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, str, this);

	if (dlg.DoModal() == IDOK)
	{
		CString strPathName = dlg.GetPathName();

		CString strStoredPath = strPathName.Left(strPathName.ReverseFind('\\'));

		// ���� ��θ� ������ ����� ���, Edit Control�� �� ����
		SetDlgItemText(IDC_EDIT_ORIGIN_FILE, strPathName);

		SetDlgItemText(IDC_EDIT_CONVERTED_FOLDER, strStoredPath);
	}
	// TODO: ���⿡ ��Ʈ�� �˸� ó���� �ڵ带 �߰��մϴ�.
}


void CTxtToXlsxDlg::OnBnClickedButtonConvert()
{
	m_ProgressCtrl.ModifyStyle(0, PBS_MARQUEE);
	m_ProgressCtrl.SetMarquee(1, 30);
	GetDlgItem(IDC_BUTTON_CONVERT)->EnableWindow(0);
	GetDlgItem(IDC_BUTTON_SELECT)->EnableWindow(0);

	m_pWaitingThread = AfxBeginThread(waiting, this);
	m_pSolveThread = AfxBeginThread(solve, this);
	
	// TODO: ���⿡ ��Ʈ�� �˸� ó���� �ڵ带 �߰��մϴ�.
}