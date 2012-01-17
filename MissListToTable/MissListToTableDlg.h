
// MissListToTableDlg.h : 头文件
//

#pragma once

#include <vector>
#include <map>
#include "afxwin.h"

// CMissListToTableDlg 对话框
class CMissListToTableDlg : public CDialogEx
{
// 构造
public:
	CMissListToTableDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_MISSLISTTOTABLE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
