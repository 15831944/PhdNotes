#include "stdafx.h"
#include "PhdLibXLExcel.h"

PhdLibXLExcel::PhdLibXLExcel()
	:m_pBook(nullptr)
	,m_pSheet(nullptr)
{

}

PhdLibXLExcel::~PhdLibXLExcel()
{
	if (m_pBook)
		m_pBook->release();
}

bool PhdLibXLExcel::CreateExcel(LPCTSTR szExcelPath)
{
	CString strSuffix = ::PathFindExtension(szExcelPath);//获取路径的扩展名
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

bool PhdLibXLExcel::OpenExcel(LPCTSTR szExcelPath)
{
	CString strSuffix = ::PathFindExtension(szExcelPath);//获取路径的扩展名
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

bool PhdLibXLExcel::OpenActiveSheet()
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

bool PhdLibXLExcel::OpenSheetByIndex(int nSheetIndex)
{
	m_pSheet = m_pBook->getSheet(nSheetIndex);
	if (!m_pSheet)
		return false;
	else
		return true;
}

bool PhdLibXLExcel::OpenSheetByName(LPCTSTR szSheetName)
{
	bool bFlag = false;
	int nCount = m_pBook->sheetCount();
	for (int i = 0; i < nCount; i++)
	{
		auto pSheet = m_pBook->getSheet(i);
		auto pName = pSheet->name();
		if (_tcsicmp(szSheetName, pName) == 0)
		{
			bFlag = true;
			m_pSheet = pSheet;
			break;
		}
	}
	return bFlag;
}

bool PhdLibXLExcel::LaunchExcel() const
{
	HINSTANCE hNewExe = ::ShellExecute(NULL, L"open", m_strExcelPath, NULL, NULL, SW_SHOW);
	if ((DWORD)hNewExe > 32)
		return true;
	else
		return false;
}

bool PhdLibXLExcel::GetCellTextOfStr(int nRowIndex, int nColIndex, CString& strText, 
	libxl::Format** format /*= 0*/) const
{
	auto text = m_pSheet->readStr(nRowIndex, nColIndex, format);
	if (!text)
	{
		auto error = m_pBook->errorMessage();
		return false;
	}
		
	strText = text;
	return true;
}

bool PhdLibXLExcel::GetCellTextOfNum(int nRowIndex, int nColIndex, double& dText, libxl::Format** format /*= 0*/) const
{
	dText = m_pSheet->readNum(nRowIndex, nColIndex, format);
	if (*format == 0)
	{
		auto error = m_pBook->errorMessage();
		return false;
	}
	return true;
}

bool PhdLibXLExcel::GetCellTextOfDate(int nRowIndex, int nColIndex, CString& strDate, libxl::Format** format /*= 0*/) const
{
	if (m_pSheet->isDate(nRowIndex, nColIndex))
		return false;

	double dTime = m_pSheet->readNum(nRowIndex, nColIndex, format);
	if (*format == 0)
	{
		auto error = m_pBook->errorMessage();
		return false;
	}

	int nYear = 0, nMonth = 0, nDay = 0, nHour = 0, nMin = 0,
		nSec = 0, nMsec = 0;
	if (!m_pBook->dateUnpack(dTime, &nYear, &nMonth, &nDay, &nHour, &nMin, &nSec, &nMsec))
		return false;
	strDate.Format(_T("%d年%2d月%2d日"), nYear, nMonth, nDay);
	strDate.Replace(_T(' '), _T('0'));

	return true;
}

bool PhdLibXLExcel::GetPrintArea(int* rowFirstIndex, int* rowLastIndex,
	int* colFirstIndex, int* colLastIndex) const
{
	return m_pSheet->printArea(rowFirstIndex, rowLastIndex, colFirstIndex, colLastIndex);
}

int PhdLibXLExcel::GetUsedRowCount() const
{
	return m_pSheet->lastRow();
}

int PhdLibXLExcel::GetUsedColCount() const
{
	return m_pSheet->lastCol();
}

libxl::Font* PhdLibXLExcel::GetFont(LPCTSTR szName) const
{
	libxl::Font* pFont = nullptr;
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

libxl::Format* PhdLibXLExcel::GetFormat(int nRowIndex, int nColIndex) const
{
	return m_pSheet->cellFormat(nRowIndex, nColIndex);
}

bool PhdLibXLExcel::SetCellTextOfStr(int nRowIndex, int nColIndex, LPCTSTR szText, 
	libxl::Format* pFormat /*= 0*/) const
{
	return m_pSheet->writeStr(nRowIndex, nColIndex, szText, pFormat);
}

bool PhdLibXLExcel::SetCellTextOfNum(int nRowIndex, int nColIndex, double dNumber, libxl::Format* pFormat /*= 0*/) const
{
	return m_pSheet->writeNum(nRowIndex, nColIndex, dNumber, pFormat);
}

bool PhdLibXLExcel::SetCellTextOfDate(int nRowIndex, int nColIndex, int nYear, int nMonth,
	int nDay, int nHour /*= 0*/, int nMin /*= 0*/, int nSec /*= 0*/, int nMsec /*= 0*/, 
	libxl::Format* pFormat /*= 0*/) const
{
	double dTime = m_pBook->datePack(nYear, nMonth, nDay, nHour, nMin, nSec, nMsec);
	
	if (pFormat == 0)
		pFormat = m_pBook->addFormat();
	pFormat->setNumFormat(NUMFORMAT_DATE);
	bool bRt = m_pSheet->writeNum(nRowIndex, nColIndex, dTime, pFormat);
	return bRt;
}

libxl::Format* PhdLibXLExcel::GetNewFormat() const
{
	if (!m_pBook)
		return NULL;
	libxl::Format* format = m_pBook->addFormat();
	if (!format)
		return NULL;

	return format;
}

void PhdLibXLExcel::SetFormatAlignCenter(libxl::Format* pFormat) const
{
	pFormat->setAlignH(ALIGNH_CENTER);
	pFormat->setAlignV(ALIGNV_CENTER);
}

bool PhdLibXLExcel::SetRowHeight(int nRowIndex, double dHeight, libxl::Format* format /*= 0*/) const
{
	return m_pSheet->setRow(nRowIndex, dHeight, format);
}

bool PhdLibXLExcel::SetColWidth(int nColFirst, int nColLast, double dWidth, libxl::Format* format /*= 0*/) const
{
	return m_pSheet->setCol(nColFirst, nColLast, dWidth, format);
}

bool PhdLibXLExcel::MergeCells(int nRowFirst, int nRowLast, int nColFirst, int nColLast) const
{
	return m_pSheet->setMerge(nRowFirst, nRowLast, nColFirst, nColLast);
}

bool PhdLibXLExcel::InsertRow(int nRowIndex) const
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
		nMaxCol = m_pSheet->lastCol();
	else
		nMaxCol = nColLast + 1;
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
		m_pSheet->setCellFormat(nRowIndex, i, pFormat);
	}
	return true;
}

bool PhdLibXLExcel::InsertRow(int nRowFirstIndex, int nRowLastIndex) const
{
	return m_pSheet->insertRow(nRowFirstIndex, nRowLastIndex);
}

void PhdLibXLExcel::SetPrintArea(int rowFirstIndex, int rowLastIndex, int colFirstIndex, int colLastIndex) const
{
	m_pSheet->setPrintArea(rowFirstIndex, rowLastIndex, colFirstIndex, colLastIndex);
}

void PhdLibXLExcel::ClearSheetContents() const
{
	m_pSheet->clear();
}
