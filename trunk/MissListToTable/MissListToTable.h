
// MissListToTable.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CMissListToTableApp:
// �йش����ʵ�֣������ MissListToTable.cpp
//

class CMissListToTableApp : public CWinApp
{
public:
	CMissListToTableApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CMissListToTableApp theApp;