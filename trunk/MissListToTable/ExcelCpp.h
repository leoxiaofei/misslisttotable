#pragma once
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"
#include "CBorders.h"

class ExcelCpp
{
public:
	ExcelCpp(void);
	~ExcelCpp(void);

	//************************************
	// Method:    Init
	// FullName:  ExcelCpp::Init
	// Access:    public 
	// Returns:   bool
	// Qualifier: ��ʼ��Excel���񣨴����ڴ���ļ���
	//************************************
	bool Init();

	//************************************
	// Method:    Init
	// FullName:  ExcelCpp::Init
	// Access:    public 
	// Returns:   bool
	// Qualifier: ��ʼ������Excel�ļ�
	// Parameter: const CString & strTemplateName
	//************************************
	bool InitByFile(const CString & strTemplateName);

	bool Save(const CString & strFileName);
	void ShowExcel();
	void OpenSheet(short sNo);
	void OpenSheet(const CString& strSheetName);

	CString GetValue();

	void SetRange(const CString& strCell);
	void SetRange(const CString& strCellBegin, const CString& strCellEnd);
	void SetValue(const CString& strValue);
	void SetValue(int nValue);
	void SetNumberFormat(const CString & strFormat);
	void SetFormula(const CString & strFormula);
	void AutoFit(bool bRow = false);
	void Merge(bool bAcross = false);
	void BorderAround(int LineStyle, int Weight, int ColorIndex = -4105);

	//************************************
	// Method:    SetHorizontalAlignment
	// FullName:  ExcelCpp::SetHorizontalAlignment
	// Access:    public 
	// Returns:   void
	// Qualifier: ˮƽ����
	// Parameter: int nType //Ĭ�ϣ�1,���У�-4108,��-4131,�ң�-4152
	//************************************
	void SetHorizontalAlignment(int nType = -4108);

	//************************************
	// Method:    SetVerticalAlignment
	// FullName:  ExcelCpp::SetVerticalAlignment
	// Access:    public 
	// Returns:   void
	// Qualifier: ��ֱ����
	// Parameter: int nType //Ĭ�ϣ�2,���У�-4108,��-4160,�ң�-4107
	//************************************
	void SetVerticalAlignment(int nType = -4108);

	//************************************
	// Method:    GetUsedMaxRowCount
	// FullName:  ExcelCpp::GetUsedMaxRowCount
	// Access:    public 
	// Returns:   int
	// Qualifier: �õ���ǰ����ʹ�õ��������
	//************************************
	int GetUsedMaxRowCount();

	//************************************
	// Method:    GetUsedMaxColumnCount
	// FullName:  ExcelCpp::GetUsedMaxColumnCount
	// Access:    public 
	// Returns:   int
	// Qualifier: �õ���ǰ����ʹ�õ��������
	//************************************
	int GetUsedMaxColumnCount();

	CString GetText();

public:
	static CString GetCellName(int nRow, int nColumn);
	static CString GetColumnName(int nColumn);

private:
	CApplication m_App;
	CWorkbooks m_Books;
	CWorkbook m_Book;
	CWorksheets m_Sheets;
	CWorksheet m_Sheet;
	CRange m_Range;
	CRange m_Cols;
	CBorders m_Borders;
	bool m_bShow;

	static COleVariant covEmpty;
};

