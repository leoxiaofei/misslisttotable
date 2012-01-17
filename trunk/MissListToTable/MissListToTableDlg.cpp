
// MissListToTableDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CMissListToTableDlg �Ի���




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


// CMissListToTableDlg ��Ϣ�������

BOOL CMissListToTableDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
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
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CMissListToTableDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
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
	if (root)//������ڵ�ApplictionConfig�Ƿ����
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
		AfxMessageBox(_T("��������Ҫ��ѯ��ֵ��"));
		return;
	}

	m_ComboCondition.GetWindowText(strCondition);

	if(strCondition.IsEmpty())
	{
		AfxMessageBox(_T("��ѡ����Ҫ��ѯ��������"));
		return;
	}

	std::map<CString,CString> m_mapData;
	ExcelCpp SrcFile,DesFile;
	if(SrcFile.InitByFile(m_strCurrentPath + _T("/") + m_strDBName))
	{
		SrcFile.OpenSheet(m_strTableName);
		int nSrcRow = SrcFile.GetUsedMaxRowCount();
		int nSrcColumn = SrcFile.GetUsedMaxColumnCount();
		
		//�ҵ�����������
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
			AfxMessageBox(_T("û���ҵ����������С�"));
			return;
		}

		//���������Ƿ����
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
			AfxMessageBox(_T("û���ҵ�����Ա��"));
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

				//������Ҫ��д������
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
	SetDlgItemText(IDC_LBL_INIT,_T("��ʼ�����"));
}
