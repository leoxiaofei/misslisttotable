
// MissListToTableDlg.h : ͷ�ļ�
//

#pragma once

#include <vector>
#include <map>
#include "afxwin.h"

// CMissListToTableDlg �Ի���
class CMissListToTableDlg : public CDialogEx
{
// ����
public:
	CMissListToTableDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_MISSLISTTOTABLE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

protected:
	void LoadItemByXml();
	void LoadTemplate();

private:
	void UTF8ToUnicode(const char* pbuffer, CString &strOut);
	//void UnicodeToUTF8(CString &strOut, const char* pbuffer);
private:
	typedef std::vector< std::pair<CString,CString> > VecIndex;
	typedef std::map< CString,std::vector<CString> > MapTemplate;
	VecIndex m_vecIndex;
	MapTemplate m_mapTemplate;
	CString m_strCurrentPath;
	CString m_strDBName;
	CString m_strTableName;
	//std::map<>
public:
	afx_msg void OnBnClickedBtnQuery();
	afx_msg void OnBnClickedBtnInit();
	CComboBox m_ComboCondition;
};
