
// TxtToXlsx.h : PROJECT_NAME ���� ���α׷��� ���� �� ��� �����Դϴ�.
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH�� ���� �� ������ �����ϱ� ���� 'stdafx.h'�� �����մϴ�."
#endif

#include "resource.h"		// �� ��ȣ�Դϴ�.


// CTxtToXlsxApp:
// �� Ŭ������ ������ ���ؼ��� TxtToXlsx.cpp�� �����Ͻʽÿ�.
//

class CTxtToXlsxApp : public CWinApp
{
public:
	CTxtToXlsxApp();

// �������Դϴ�.
public:
	virtual BOOL InitInstance();

// �����Դϴ�.

	DECLARE_MESSAGE_MAP()
};

extern CTxtToXlsxApp theApp;