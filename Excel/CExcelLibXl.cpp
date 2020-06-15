#include "StdAfx.h"
#include "CExcelLibXl.h"
#include "../NameSpace/PhdMethod.h"
#include "../NameSpace/PhdConver.h"


CExcelLibXl::CExcelLibXl()
	:m_pBook(NULL)
	, m_pSheet(NULL)
{

}


CExcelLibXl::~CExcelLibXl()
{
	if (m_pBook)
		m_pBook->release();
}

bool CExcelLibXl::CreateExcelFile(LPCTSTR szExcelPath)
{
	CString strSuffix = GetSuffix(szExcelPath);
	if (_tcsicmp(strSuffix, _T(".xls")) == 0)
		m_pBook = xlCreateBook();
	else
		m_pBook = xlCreateXMLBook();
	if (!m_pBook)
		return false;
	//破解
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	m_pBook->setKey(x, y);
	//
	m_pSheet = m_pBook->addSheet(_T("Sheet1"));
	if (!m_pSheet)
		return false;
	bool bRt = m_pBook->save(szExcelPath);
	m_strExcelPath = szExcelPath;
	return bRt;
}

bool CExcelLibXl::OpenExcel(LPCTSTR szExcelPath)
{
	CString strSuffix = GetSuffix(szExcelPath);
	if (_tcsicmp(strSuffix, _T(".xls")) == 0)
		m_pBook = xlCreateBook();
	else
		m_pBook = xlCreateXMLBook();
	if (!m_pBook)
		return false;
	//破解
	const wchar_t * x = L"Halil Kural";
	const wchar_t * y = L"windows-2723210a07c4e90162b26966a8jcdboe";
	m_pBook->setKey(x, y);
	//
	bool bRt = m_pBook->load(szExcelPath);
	m_strExcelPath = szExcelPath;
	return bRt;
}

bool CExcelLibXl::OpenActiveSheet()
{
	if (!m_pBook)
		return false;
	int nSheetIndex = m_pBook->activeSheet();
	if (nSheetIndex < 0)
		return false;
	m_pSheet = m_pBook->getSheet(nSheetIndex);
	if (!m_pSheet)
		return false;
	else
		return true;
}

bool CExcelLibXl::Save()
{
	if (::PathFileExists(m_strExcelPath))//如果文件存在
		PhdMethod::DeleteFileOrDir(m_strExcelPath);//删除原文件
	return m_pBook->save(m_strExcelPath);
}

bool CExcelLibXl::GetPrintArea(int* rowFirstIndex, int* rowLastIndex, 
	int* colFirstIndex, int* colLastIndex) const
{
	return m_pSheet->printArea(rowFirstIndex, rowLastIndex, colFirstIndex, colLastIndex);
}

CString CExcelLibXl::GetText(int nRowIndex, int nColIndex,int nprec/*= 2*/) const
{
	CString strText;
	if (m_pSheet->isDate(nRowIndex,nColIndex))
	{
		double dTime = m_pSheet->readNum(nRowIndex, nColIndex);
		int nYear = 0, nMonth = 0, nDay = 0, nHour = 0, nMin = 0,
			nSec = 0, nMsec = 0;
		if (!m_pBook->dateUnpack(dTime, &nYear, &nMonth, &nDay, &nHour, &nMin, &nSec, &nMsec))
			return _T("");
		strText.Format(_T("%d年%2d月%2d日"), nYear, nMonth, nDay);
		strText.Replace(_T(' '), _T('0'));
		return strText;
	}

	//
	LPCTSTR szText = m_pSheet->readStr(nRowIndex, nColIndex);
	CString strError = m_pBook->errorMessage();
	if (strError != OK)
	{
		double dText = m_pSheet->readNum(nRowIndex, nColIndex);
		strError = m_pBook->errorMessage();
		if (strError != OK)
			return _T("");
		strText = PhdConver::DoubleToStr(dText, nprec);
	}
	else
	{
 		strText = szText;
	}
		
	return strText;
}

CString CExcelLibXl::GetTextString(int nRowIndex, int nColIndex) const
{
	LPCTSTR szText = m_pSheet->readStr(nRowIndex, nColIndex);
	if (!szText)
	{
		CString strError = m_pBook->errorMessage();
		return strError;
	}
		
	CString strText = szText;
	return strText;
}

double CExcelLibXl::GetTextNumber(int nRowIndex, int nColIndex) const
{
	double dText = m_pSheet->readNum(nRowIndex, nColIndex);
	return dText;
}

CString CExcelLibXl::GetTextTime(int nRowIndex, int nColIndex) const
{
	CString strText;
	double dTime = m_pSheet->readNum(nRowIndex, nColIndex);
	int nYear = 0, nMonth = 0, nDay = 0, nHour = 0, nMin = 0,
		nSec = 0, nMsec = 0;
	if (!m_pBook->dateUnpack(dTime, &nYear, &nMonth, &nDay, &nHour, &nMin, &nSec, &nMsec))
		return _T("");
	strText.Format(_T("%d年%2d月%2d日"), nYear, nMonth, nDay);
	strText.Replace(_T(' '), _T('0'));
	return strText;
}

bool CExcelLibXl::LaunchExcel() const
{
	HINSTANCE hNewExe = ::ShellExecute(NULL, L"open", m_strExcelPath, NULL, NULL, SW_SHOW);
	if ((DWORD)hNewExe > 32)
		return true;
	else
		return false;
}

void CExcelLibXl::SetFormatAlignCenter(libxl::Format* pFormat) const
{
	pFormat->setAlignH(ALIGNH_CENTER);
	pFormat->setAlignV(ALIGNV_CENTER);
}

bool CExcelLibXl::SetText(int nRowIndex, int nColIndex, LPCTSTR szText, libxl::Format* pFormat /*= NULL*/) const
{
	CString strText = szText;
	if (strText == _T(""))
		return true;

	bool bRt = StringIsNumber(szText);
	if (!bRt)
		bRt = m_pSheet->writeStr(nRowIndex, nColIndex, szText, pFormat);
	else
	{
		double dNumber = _wtof(szText);
		bRt = m_pSheet->writeNum(nRowIndex, nColIndex, dNumber, pFormat);
	}

	return bRt;
}

bool CExcelLibXl::SetTextString(int nRowIndex, int nColIndex, LPCTSTR szText,
	libxl::Format* pFormat /*= NULL*/) const
{
	return m_pSheet->writeStr(nRowIndex, nColIndex, szText, pFormat);
}

bool CExcelLibXl::SetTextNumber(int nRowIndex, int nColIndex, 
	double dNumber, libxl::Format* pFormat /*= NULL*/) const
{
	return m_pSheet->writeNum(nRowIndex, nColIndex, dNumber, pFormat);
}

bool CExcelLibXl::SetTextTime(int nRowIndex, int nColIndex, int nYear, int nMonth, int nDay, int nHour /*= 0*/, int nMin /*= 0*/, int nSec /*= 0*/, int nMsec /*= 0*/, libxl::Format* pFormat /*= NULL*/) const
{
	double dTime = m_pBook->datePack(nYear, nMonth, nDay, nHour, nMin, nSec, nMsec);
	return m_pSheet->writeNum(nRowIndex, nColIndex, dTime, pFormat);
}

int CExcelLibXl::GetFirstRow() const
{
	if (!m_pSheet)
		return -1;
	return m_pSheet->firstRow();
}

int CExcelLibXl::GetLastRow() const
{
	if (!m_pSheet)
		return -1;
	return m_pSheet->lastRow();
}

int CExcelLibXl::GetFirstCol() const
{
	if (!m_pSheet)
		return -1;
	return m_pSheet->firstCol();
}

int CExcelLibXl::GetLastCol() const
{
	if (!m_pSheet)
		return -1;
	return m_pSheet->lastCol();
}

bool CExcelLibXl::SetColWidth(int colFirst, int colLast, double width,
	libxl::Format* pFormat /*= NULL*/) const
{
	bool bRt = m_pSheet->setCol(colFirst, colLast, width, pFormat);
	CString strError = m_pBook->errorMessage();
	return bRt;
}

bool CExcelLibXl::SetRowHeight(int nRowIndex, double dHeight, libxl::Format* pFormat /*= NULL*/) const
{
	return m_pSheet->setRow(nRowIndex, dHeight, pFormat);
}

bool CExcelLibXl::SetColFormat(int colFirst, int colLast,
	libxl::Format* pFormat) const
{
	if (colLast < colFirst)
		return false;

	double dWidth = 0;
	if (colFirst == colLast)
	{
		dWidth = m_pSheet->colWidth(colFirst);
		return m_pSheet->setCol(colFirst, colLast, dWidth, pFormat);
	}
	else
	{
		for (int i = colFirst; i < colLast; i++)
		{
			dWidth = m_pSheet->colWidth(i);
			m_pSheet->setCol(i, i, dWidth, pFormat);
		}
		return true;
	}
}

bool CExcelLibXl::SetRowFormat(int rowIndex, libxl::Format* pFormat) const
{
	double dHeight = m_pSheet->rowHeight(rowIndex);
	return m_pSheet->setRow(rowIndex, dHeight, pFormat);
}

bool CExcelLibXl::MergeCells(int rowFirst, int rowLast, int colFirst, int colLast) const
{
	return m_pSheet->setMerge(rowFirst,rowLast,colFirst,colLast);
}

bool CExcelLibXl::InsertRow(int nRowFirstIndex, int nRowLastIndex) const
{
	return m_pSheet->insertRow(nRowFirstIndex, nRowLastIndex);
}

bool CExcelLibXl::InsertRow(int nRowIndex) const
{
	//设置打印区域
	int nRowFirst = 0, nRowLast = 0, nColFirst = 0, nColLast = 0;
	GetPrintArea(&nRowFirst, &nRowLast, &nColFirst, &nColLast);
	if (nRowIndex <= nRowLast)
		m_pSheet->setPrintArea(nRowFirst, nRowLast + 1, nColFirst, nColLast);
	//插入新的一行
	bool bRt = m_pSheet->insertRow(nRowIndex, nRowIndex);
	if (!bRt)
		return false;
	//最大列
	int nMaxCol = 0;
	if (nColLast == 0)
		nMaxCol = GetLastCol();
	else
		nMaxCol = nColLast+1;
	//模板行
	int nRowTemplate = 0;	
	if (nRowIndex == 0)
		nRowTemplate = 1;
	else
		nRowTemplate = nRowIndex - 1;
	//设置行高
	double dHeight = m_pSheet->rowHeight(nRowTemplate);	
	SetRowHeight(nRowIndex, dHeight);
	//设置单元格格式
	libxl::Format* pFormat = NULL;
	for (int i = 0; i < nMaxCol; i++)
	{
		pFormat = m_pSheet->cellFormat(nRowTemplate, i);
		m_pSheet->setCellFormat(nRowIndex,i,pFormat);
	}
	return true;
}

bool CExcelLibXl::SetPrintArea(int rowFirstIndex, int rowLastIndex, int colFirstIndex, int colLastIndex) const
{
	m_pSheet->setPrintArea(rowFirstIndex, rowLastIndex, colFirstIndex, colLastIndex);
	return true;
}

void CExcelLibXl::ClearSheetContents() const
{
	m_pSheet->clear();
	m_pBook->save(m_strExcelPath);
}

CString CExcelLibXl::GetSuffix(LPCTSTR szExcelPath) const
{
	CString strPath = szExcelPath;
	int nFindIndex = strPath.ReverseFind(_T('.'));
	if (nFindIndex == -1)
		return _T("");
	CString strSuffix = strPath.Right(strPath.GetLength() - nFindIndex);
	return strSuffix;
}

bool CExcelLibXl::StringIsNumber(LPCTSTR szString) const
{
	CString strTemp = szString;
	bool bFlag = true;
	for (int i = 0; i < strTemp.GetLength(); i++)
	{
		TCHAR str = strTemp[i];
		if (str != _T('.') && ((str < _T('0') || str > _T('9'))) && str != _T('-'))
		{
			bFlag = false;
		}
	}
	return bFlag;
}

libxl::Format* CExcelLibXl::AddFormat() const
{
	if (!m_pBook)
		return NULL;
	libxl::Format* format = m_pBook->addFormat();
	if (!format)
		return NULL;

	return format;
}

libxl::Font* CExcelLibXl::GetFont(LPCTSTR szName) const
{
	libxl::Font* pFont = NULL;
	int nCount = m_pBook->fontSize();
	if (nCount > 0)
	{
		for (int i = 0; i < nCount; i++)
		{
			libxl::Font* pFontTemp = m_pBook->font(i);
			if (!pFontTemp)
				continue;
			CString strName = pFontTemp->name();
			if (_tcscmp(strName, szName) == 0)
			{
				pFont = pFontTemp;
				break;
			}
		}
	}
	if (!pFont)
	{
		pFont = m_pBook->addFont();
		pFont->setName(szName);
	}
	return pFont;
}

libxl::Format* CExcelLibXl::GetFormat(int nRowIndex, int nColIndex) const
{
	return m_pSheet->cellFormat(nRowIndex, nColIndex);
}
