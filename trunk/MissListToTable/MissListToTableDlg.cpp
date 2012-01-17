
// MissListToTableDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "MissListToTable.h"
#include "MissListToTableDlg.h"
#include "afxdialogex.h"
#include "tinyxml.h"
#include "ExcelCpp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMissListToTableDlg 对话框




CMissListToTableDlg::CMissListToTableDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMissListToTableDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMissListToTableDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO_CONDITION, m_ComboCondition);
}

BEGIN_MESSAGE_MAP(CMissListToTableDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_QUERY, &CMissListToTableDlg::OnBnClickedBtnQuery)
	ON_BN_CLICKED(IDC_BTN_INIT, &CMissListToTableDlg::OnBnClickedBtnInit)
END_MESSAGE_MAP()


// CMissListToTableDlg 消息处理程序

BOOL CMissListToTableDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	TCHAR pChar[MAX_PATH];
	GetModuleFileName(NULL,pChar,MAX_PATH);
	int nIndex(0),nFind(0);
	while( pChar[nIndex] != '\0' )
	{
		if(pChar[nIndex] == '\\' )
		{
			nFind = nIndex;
		}
		++nIndex;
	}
	if(nFind != 0)
	{
		pChar[nFind] = '\0';
	}
	SetCurrentDirectory(pChar);
	m_strCurrentPath = pChar;
	OnBnClickedBtnInit();
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CMissListToTableDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CMissListToTableDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CMissListToTableDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CMissListToTableDlg::LoadItemByXml()
{
	std::vector<std::pair<CString,CString>>().swap(m_vecIndex);
	//m_ComboCondition.Clear();
	m_ComboCondition.ResetContent();
	TiXmlDocument doc("Config.xml");
	doc.LoadFile();
	TiXmlElement* root = doc.FirstChildElement("ApplictionConfig");
	const char* CurStr;
	if (root)//检测主节点ApplictionConfig是否存在
	{
		TiXmlElement * CurElement = root->FirstChildElement("DataSource");
		if (CurElement)
		{
			CurStr = CurElement->Attribute("DBName");
			UTF8ToUnicode(CurStr,m_strDBName);

			CurStr = CurElement->Attribute("TableName");
			UTF8ToUnicode(CurStr,m_strTableName);
		}


		CurElement = root->FirstChildElement("Conditions");
		CurElement = CurElement->FirstChildElement();
		
		CString strCName;
		while(CurElement)
		{
			CurStr = CurElement->Attribute("CName");
			UTF8ToUnicode(CurStr,strCName);

			m_ComboCondition.AddString(strCName);

			CurElement=CurElement->NextSiblingElement();
		}
		if (m_ComboCondition.GetCount() > 0)
		{
			m_ComboCondition.SetCurSel(0);
		}
		CurElement = root->FirstChildElement("Items");
		CurElement = CurElement->FirstChildElement();
		std::pair<CString,CString> pa;
		while(CurElement)
		{
			CurStr = CurElement->Attribute("Src");
			UTF8ToUnicode(CurStr,pa.first);

			CurStr = CurElement->Attribute("Des");
			UTF8ToUnicode(CurStr,pa.second);

			m_vecIndex.push_back(pa);

			CurElement=CurElement->NextSiblingElement();
		}
	}
}

void CMissListToTableDlg::UTF8ToUnicode( const char* pIn, CString &strOut )
{
	int nSize = MultiByteToWideChar( CP_UTF8, 0, pIn, -1, 0, 0 );
	wchar_t* pbuffer = new wchar_t[nSize];
	memset( pbuffer, 0, nSize * sizeof(wchar_t) );
	MultiByteToWideChar( CP_UTF8, 0, pIn, -1, pbuffer, nSize );
	strOut = pbuffer;
	delete[] pbuffer;
	 
}



void CMissListToTableDlg::OnBnClickedBtnQuery()
{
	// TODO: Add your control notification handler code here
	CString strName,strCondition;
	GetDlgItemText(IDC_EDIT_NAME,strName);

	if(strName.IsEmpty())
	{
		AfxMessageBox(_T("请输入所要查询的值。"));
		return;
	}

	m_ComboCondition.GetWindowText(strCondition);

	if(strCondition.IsEmpty())
	{
		AfxMessageBox(_T("请选择所要查询的条件。"));
		return;
	}

	std::map<CString,CString> m_mapData;
	ExcelCpp SrcFile,DesFile;
	if(SrcFile.InitByFile(m_strCurrentPath + _T("/") + m_strDBName))
	{
		SrcFile.OpenSheet(m_strTableName);
		int nSrcRow = SrcFile.GetUsedMaxRowCount();
		int nSrcColumn = SrcFile.GetUsedMaxColumnCount();
		
		//找到姓名所在列
		int nColumnName(-1);
		for (int ix = 1; ix <= nSrcColumn; ++ix)
		{
			SrcFile.SetRange(ExcelCpp::GetCellName(1,ix));
			if(strCondition == SrcFile.GetText())
			{
				nColumnName = ix;
				break;
			}
		}

		if(nColumnName == -1)
		{
			AfxMessageBox(_T("没有找到姓名所在列。"));
			return;
		}

		//查找姓名是否存在
		int nRowName(-1);
		for(int ix = 2; ix <= nSrcRow; ++ix)
		{
			SrcFile.SetRange(ExcelCpp::GetCellName(ix,nColumnName));
			if(strName == SrcFile.GetText())
			{
				nRowName = ix;
				break;
			}
		}

		if(nRowName == -1)
		{
			AfxMessageBox(_T("没有找到该人员。"));
			return;
		}
		
		CString strFile;
		for (int ix = 1; ix <= nSrcColumn; ++ix)
		{
			SrcFile.SetRange(ExcelCpp::GetCellName(1,ix));
			strFile = SrcFile.GetText();
			for(std::vector<std::pair<CString,CString>>::iterator itor = m_vecIndex.begin();
				itor != m_vecIndex.end(); ++itor)
			{
				if(strFile == itor->second)
				{
					SrcFile.SetRange(ExcelCpp::GetCellName(nRowName,ix));
					m_mapData[itor->first] = SrcFile.GetText();
					break;
				}
			}
		}
	}



	if(DesFile.InitByFile(m_strCurrentPath + _T("/Template.xlt")))
	{
		std::map<CString,CString>::iterator itorFind;
		DesFile.OpenSheet(1);
		for(MapTemplate::iterator itor = m_mapTemplate.begin(); 
			itor != m_mapTemplate.end(); ++itor)
		{
			for(int ix = 0; ix != itor->second.size(); ++ix)
			{
				DesFile.SetRange(itor->second[ix]);

				//查找所要填写的内容
				itorFind = m_mapData.find(itor->first);
				if(itorFind != m_mapData.end())
				{
					DesFile.SetValue(itorFind->second);
				}
			}
		}
		DesFile.ShowExcel();
	}

}


void CMissListToTableDlg::OnBnClickedBtnInit()
{
	// TODO: Add your control notification handler code here
	LoadItemByXml();
	LoadTemplate();
}

void CMissListToTableDlg::LoadTemplate()
{
	m_mapTemplate.clear();
	ExcelCpp TempFile;
	if(TempFile.InitByFile(m_strCurrentPath + _T("/Template.xlt")))
	{
		CString strValue;
		TempFile.OpenSheet(1);
		int nRow = TempFile.GetUsedMaxRowCount();
		int nColumn = TempFile.GetUsedMaxColumnCount();
		for(int i = 1; i<= nRow; ++i)
		{
			for(int j = 1; j <= nColumn; ++j)
			{
				TempFile.SetRange(ExcelCpp::GetCellName(i,j));
				strValue = TempFile.GetText();
				for(std::vector<std::pair<CString,CString>>::iterator itor = m_vecIndex.begin();
					itor != m_vecIndex.end(); ++itor)
				{
					if(strValue == itor->first)
					{
						m_mapTemplate[strValue].push_back(ExcelCpp::GetCellName(i,j));
						break;
					}
				}
			}
		}
	}
	SetDlgItemText(IDC_LBL_INIT,_T("初始化完成"));
}
