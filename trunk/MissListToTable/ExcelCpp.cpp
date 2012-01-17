#include "StdAfx.h"
#include "ExcelCpp.h"


ExcelCpp::ExcelCpp(void):
	m_bShow(false)
{

}


ExcelCpp::~ExcelCpp(void)
{
	m_Range.ReleaseDispatch();
	m_Cols.ReleaseDispatch();
	m_Sheet.ReleaseDispatch();
	m_Sheets.ReleaseDispatch();
	m_Book.ReleaseDispatch();
	m_Books.ReleaseDispatch();
	if ( !m_bShow )
	{
		m_App.Quit();
	}
	m_App.ReleaseDispatch();
}

bool ExcelCpp::Init()
{
	if(m_App.CreateDispatch(_T("Excel.Application")))
	{
		m_App.put_AlertBeforeOverwriting(FALSE);
		m_App.put_DisplayAlerts(FALSE);
		m_Books=m_App.get_Workbooks();
		m_Book=m_Books.Add(covEmpty);
		m_Sheets=m_Book.get_Sheets();
		return true;
	}
	return false;
}

bool ExcelCpp::InitByFile( const CString & strTemplateName )
{
	if(m_App.CreateDispatch(_T("Excel.Application")))
	{
		m_App.put_AlertBeforeOverwriting(FALSE);
		m_App.put_DisplayAlerts(FALSE);

		m_Books=m_App.get_Workbooks();
		m_Book=m_Books.Add(COleVariant(strTemplateName));
		m_Sheets=m_Book.get_Sheets();
		return true;
	}
	return false;
}

bool ExcelCpp::Save( const CString & strFileName )
{
	m_Book.SaveAs(
		COleVariant(strFileName),
		covEmpty,
		covEmpty,
		covEmpty,
		covEmpty,
		covEmpty,
		(long)0,
		covEmpty,
		covEmpty,
		covEmpty,
		covEmpty,
		covEmpty);
	return true;
}

CString ExcelCpp::GetCellName( int nRow, int nColumn )
{
	ASSERT(nRow>0);
	ASSERT(nColumn>0);
	ASSERT(nColumn<256);
	CString strRet;
	if (nColumn > 26)
	{
		strRet.AppendChar('A' + (nColumn-1) / 26 - 1);
	}
	strRet.AppendChar('A' + (nColumn-1) % 26);
	strRet.AppendFormat(_T("%d"),nRow);
	return strRet;
}

CString ExcelCpp::GetColumnName( int nColumn )
{
	ASSERT(nColumn>0);
	ASSERT(nColumn<256);
	CString strRet;
	if (nColumn > 26)
	{
		strRet.AppendChar('A' + nColumn / 26 - 1);
	}
	strRet.AppendChar('A' + (nColumn - 1) % 26 );
	return strRet;
}

void ExcelCpp::SetRange( const CString& strCell )
{
	m_Range = m_Sheet.get_Range(COleVariant(strCell),COleVariant(strCell));
}

void ExcelCpp::SetRange( const CString& strCellBegin, const CString& strCellEnd )
{
	m_Range = m_Sheet.get_Range(COleVariant(strCellBegin),COleVariant(strCellEnd));
}

void ExcelCpp::SetValue( const CString & strValue )
{
	m_Range.put_Value2(COleVariant(strValue));
}

void ExcelCpp::SetValue( int nValue )
{
	m_Range.put_Value2(COleVariant(static_cast<long>(nValue)));
}

void ExcelCpp::AutoFit(bool bRow)
{
	m_Cols = bRow?m_Range.get_EntireRow():m_Range.get_EntireColumn();
	m_Cols.AutoFit();
}

void ExcelCpp::SetNumberFormat( const CString & strFormat )
{
	m_Range.put_NumberFormat(COleVariant(strFormat));
}

void ExcelCpp::SetFormula( const CString & strFormula )
{
	m_Range.put_Formula(COleVariant(strFormula));
}

void ExcelCpp::Merge( bool bAcross )
{
	m_Range.Merge(_variant_t(bAcross));
	m_Range.put_HorizontalAlignment(_variant_t(3));
}

void ExcelCpp::ShowExcel()
{
	m_bShow = true;
	m_App.put_Visible(TRUE);
	m_App.put_UserControl(TRUE);
}

void ExcelCpp::BorderAround(int LineStyle, int Weight, int ColorIndex)
{
	//设置边框 
	m_Borders = m_Range.get_Borders();
	m_Borders.put_LineStyle(_variant_t((long)LineStyle));
	m_Borders.put_Weight(_variant_t((long)Weight));
	m_Borders.put_ColorIndex(_variant_t((long)ColorIndex));
// 	m_Range.BorderAround(
// 		_variant_t((long)LineStyle),
// 		_variant_t((long)Weight),
// 		_variant_t((long)ColorIndex),
// 		vtMissing);
}

//水平对齐：默认＝1,居中＝-4108,左＝-4131,右＝-4152 
void ExcelCpp::SetHorizontalAlignment( int nType )
{ 
	m_Range.put_HorizontalAlignment(_variant_t((long)nType)); 
}

void ExcelCpp::SetVerticalAlignment( int nType )
{
	//垂直对齐：默认＝2,居中＝-4108,左＝-4160,右＝-4107
	m_Range.put_VerticalAlignment(_variant_t((long)nType)); 
}

void ExcelCpp::OpenSheet( short sNo )
{
	m_Sheet = m_Sheets.get_Item(COleVariant(sNo));
}

void ExcelCpp::OpenSheet( const CString& strSheetName )
{
	m_Sheet = m_Sheets.get_Item(COleVariant(strSheetName));
}

int ExcelCpp::GetUsedMaxRowCount()
{
	CRange range;
	range = m_Sheet.get_UsedRange();  //获得Worksheet已使用的范围
	range = range.get_Rows();         //获得总行数（LPDISPATCH类型）
	return range.get_Count();         //即可获得已使用的行数了。 
}

int ExcelCpp::GetUsedMaxColumnCount()
{
	CRange range;
	range = m_Sheet.get_UsedRange();  //获得Worksheet已使用的范围
	range = range.get_Columns();      //获得总列数（LPDISPATCH类型）
	return range.get_Count();         //即可获得已使用的列数了。
}

CString ExcelCpp::GetValue()
{
	return m_Range.get_Value2().bstrVal;
}

CString ExcelCpp::GetText()
{
	return m_Range.get_Text().bstrVal;
}


COleVariant ExcelCpp::covEmpty((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
