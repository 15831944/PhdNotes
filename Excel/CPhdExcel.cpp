#include "StdAfx.h"
#include "CPhdExcel.h"

CPhdExcel::CPhdExcel()
	:m_app(nullptr)
	,m_books(nullptr)
	,m_book(nullptr)
	,m_sheets(nullptr)
	,m_sheet(nullptr)
	,m_range(nullptr)
{

}

CPhdExcel::~CPhdExcel()
{
	if (m_bExcelIsOpen)
		Clear();
	else
		Quit();
}

bool CPhdExcel::CreateExcel(LPCTSTR szExcelPath)
{
	CoInitialize(NULL);

	if (!m_app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("����Excelʧ��!"));
		exit(1);
		return false;
	}

	try
	{
		m_app.SetVisible(false);
		//����ǰ�����ϴ򿪵�����excel��������m_books
		m_books.AttachDispatch(m_app.GetWorkbooks(), true);
		//�������Ĺ�����
		m_book.AttachDispatch(m_books.Add(vtMissing));
		//��m_sheets
		m_sheets.AttachDispatch(m_book.GetWorksheets(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}

	//����Ϊ
	SaveAs(szExcelPath);

	return true;
}

bool CPhdExcel::CreateExcelByTemplate(LPCTSTR szTemplatePath, LPCTSTR szExcelPath)
{
	CoInitialize(NULL);

	if (!m_app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("����Excelʧ��!"));
		exit(1);
		return false;
	}

	try
	{
		m_app.SetVisible(false);
		//����ǰ�����ϴ򿪵�����excel��������m_books
		m_books.AttachDispatch(m_app.GetWorkbooks(), true);
		//�������Ĺ�����
		m_book.AttachDispatch(m_books.Add(_variant_t(szTemplatePath)));
		//��m_sheets
		m_sheets.AttachDispatch(m_book.GetWorksheets(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}

	//����Ϊ
	SaveAs(szExcelPath);

	return true;
}

bool CPhdExcel::Open(LPCTSTR szExcelPath)
{
	CoInitialize(NULL);

	if (!m_app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("����Excelʧ��!"));
		exit(1);
		return false;
	}
	
	try
	{
		m_app.SetVisible(false);
		//����ǰ�����ϴ򿪵�����excel��������m_books
		m_books.AttachDispatch(m_app.GetWorkbooks(), true);
		//��ָ���Ĺ�����
		m_book = m_books.Open(szExcelPath, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing);
		//��m_sheets
		m_sheets.AttachDispatch(m_book.GetWorksheets(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}

	//��m_sheet��m_range
	if (!OpenActiveSheet())
		return false;

	m_bExcelIsOpen = false;
	return true;
}

bool CPhdExcel::OpenTheAlreadyOpenExcel(LPCTSTR szExcelPath)
{
	//��ʼ��com�⣬�Ե��̵߳ķ�ʽ����com����
	CoInitialize(NULL);

	LPDISPATCH lpDisp = NULL;
	CLSID clsid;
	HRESULT hr;
	hr = ::CLSIDFromProgID(_T("Excel.Application"), &clsid); //ͨ��ProgIDȡ��CLSID
	if (FAILED(hr))
		return FALSE;//������û��excel����

	IUnknown *pUnknown = NULL;
	IDispatch *pDispatch = NULL;
	hr = ::GetActiveObject(clsid, NULL, &pUnknown); //�����Ƿ��г���������
	if (FAILED(hr))
		return false;//û��excel����������
	else
	{//��excel����������
		try
		{
			hr = pUnknown->QueryInterface(IID_IDispatch, (LPVOID *)&m_app);
			if (FAILED(hr))
				throw(_T("û��ȡ��IDispatchPtr"));
			pUnknown->Release();
			pUnknown = NULL;

			lpDisp = m_app.GetWorkbooks();
			m_books.AttachDispatch(lpDisp, TRUE);
			int nLen = m_books.GetCount();
			CString strName;
			bool bIsFind = false;
			for (int i = 1; i <= nLen; i++)
			{
				if (m_book)
					m_book.ReleaseDispatch();
				lpDisp = m_books.GetItem(_variant_t(i));
				m_book.AttachDispatch(lpDisp, TRUE);
				strName = m_book.GetFullName();

				if (_tcscmp(strName, szExcelPath) == 0)
				{
					bIsFind = true;
					break;
				}
			}
			if (!bIsFind)
			{
				m_book.ReleaseDispatch();
				m_books.ReleaseDispatch();
				m_app.ReleaseDispatch();
				CoUninitialize();//�رյ�ǰ�߳��ϵ�com�⣬ж�ظ��̼߳��ص�����dll,�ͷŸ��߳�ά��������������Դ����ǿ�ƹرո��߳��ϵ�����RPC���ӡ�
				return false;//û�ҵ�
			}

			//��m_sheets
			m_sheets.AttachDispatch(m_book.GetWorksheets(), true);
		}
		catch (CException* e)
		{
			TCHAR szError[1024];
			e->GetErrorMessage(szError, 1024);
			return false;
		}
	}

	//��m_sheet��m_range
	if (!OpenActiveSheet())
		return false;

	m_bExcelIsOpen = true;
	return true;
}

bool CPhdExcel::AutoOpen(LPCTSTR szExcelPath)
{
	if (!IsOpen(szExcelPath))
	{
		return Open(szExcelPath);
	}
	else
	{
		return OpenTheAlreadyOpenExcel(szExcelPath);
	}
}

bool CPhdExcel::IsOpen(LPCTSTR szExcelPath)
{
	LPDISPATCH lpDisp = NULL;
	CoInitialize(NULL);
	CLSID clsid;
	HRESULT hr;
	hr = ::CLSIDFromProgID(_T("Excel.Application"), &clsid); //ͨ��ProgIDȡ��CLSID
	if (FAILED(hr))
	{
		return FALSE;//������û��excel����
	}

	IUnknown *pUnknown = NULL;
	IDispatch *pDispatch = NULL;
	hr = ::GetActiveObject(clsid, NULL, &pUnknown); //�����Ƿ��г���������
	if (FAILED(hr))
	{//û��excel����������
		return false;
	}
	else
	{//��excel����������
		try
		{
			_ExApplication appTemp;
			hr = pUnknown->QueryInterface(IID_IDispatch, (LPVOID *)&appTemp);
			if (FAILED(hr))
				throw(_T("û��ȡ��IDispatchPtr"));
			pUnknown->Release();
			pUnknown = NULL;

			lpDisp = appTemp.GetWorkbooks();
			Workbooks booksTemp;
			booksTemp.AttachDispatch(lpDisp, TRUE);
			int nLen = booksTemp.GetCount();
			CString strName;
			bool bIsFind = false;
			_Workbook bookTemp;
			for (int i = 1; i <= nLen; i++)
			{
				if (bookTemp)
					bookTemp.ReleaseDispatch();
				lpDisp = booksTemp.GetItem(_variant_t(i));
				bookTemp.AttachDispatch(lpDisp, TRUE);
				strName = bookTemp.GetFullName();

				if (_tcscmp(strName, szExcelPath) == 0)
				{
					bIsFind = true;
					break;
				}
			}
			//
			bookTemp.ReleaseDispatch();
			booksTemp.ReleaseDispatch();
			appTemp.ReleaseDispatch();
			//�رյ�ǰ�߳��ϵ�com�⣬ж�ظ��̼߳��ص�����dll,�ͷŸ��߳�ά��������������Դ����ǿ�ƹرո��߳��ϵ�����RPC���ӡ�
			CoUninitialize();
			//
			if (!bIsFind)
				return false;//û�ҵ�
			else
				return true;
		}
		catch (CException* e)
		{
			TCHAR szError[1024];
			e->GetErrorMessage(szError, 1024);
			return false;
		}
	}
}

bool CPhdExcel::OpenActiveSheet()
{
	try
	{
		m_sheet.AttachDispatch(m_book.GetActiveSheet(), true);
		m_range.AttachDispatch(m_sheet.GetCells(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	
	return true;
}

bool CPhdExcel::OpenSheet(LPCTSTR szSheetName)
{
	try
	{
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		m_sheet.AttachDispatch(m_sheets.GetItem(_variant_t(szSheetName)), true);
		if (m_range)
			m_range.ReleaseDispatch();
		m_range.AttachDispatch(m_sheet.GetCells(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}

	return true;
}

bool CPhdExcel::OpenSheet(int nIndex)
{
	try
	{
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		m_sheet.AttachDispatch(m_sheets.GetItem(COleVariant((long)nIndex)), true);
		if (m_range)
			m_range.ReleaseDispatch();
		m_range.AttachDispatch(m_sheet.GetCells(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}

	return true;
}

bool CPhdExcel::Save()
{
	try
	{
		//���ò���ʾ�Ƿ񸲸Ǿ���
		m_app.SetAlertBeforeOverwriting(FALSE);
		//���ò���ʾ����
		m_app.SetDisplayAlerts(FALSE);
		//���湤����
		m_book.Save();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SaveAs(LPCTSTR szExcelPath)
{
	try
	{
		//���ò���ʾ�Ƿ񸲸Ǿ���
		m_app.SetAlertBeforeOverwriting(FALSE);
		//���ò���ʾ����
		m_app.SetDisplayAlerts(FALSE);
		//����Ϊ������
		m_book.SaveAs(_variant_t(szExcelPath), vtMissing, vtMissing, vtMissing, vtMissing, vtMissing
			, 0, vtMissing, vtMissing, vtMissing, vtMissing);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::Clear()
{
	try
	{
		if (m_range)
			m_range.ReleaseDispatch();
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		if (m_sheets)
			m_sheets.ReleaseDispatch();
		if (m_book)
			m_book.ReleaseDispatch();
		if (m_books)
			m_books.ReleaseDispatch();
		if (m_app)
			m_app.ReleaseDispatch();

		CoUninitialize();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::Quit()
{
	if (!m_app)
		return false;

	try
	{
		m_app.SetVisible(FALSE);
		m_app.SetDisplayAlerts(FALSE);

		if (m_books)
			m_books.Close();	//�رչ���������

		m_app.Quit();		//�˳�excelӦ�ó���

		if (m_range)
			m_range.ReleaseDispatch();
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		if (m_sheets)
			m_sheets.ReleaseDispatch();
		if (m_book)
			m_book.ReleaseDispatch();
		if (m_books)
			m_books.ReleaseDispatch();
		if (m_app)
			m_app.ReleaseDispatch();

		CoUninitialize();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

int CPhdExcel::GetSheetCount() 
{
	long lCount = 0;
	try
	{
		lCount = m_sheets.GetCount();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return 0;
	}
	return lCount;
}

CString CPhdExcel::GetSheetNameByIndex(int nIndex)
{
	CString strSheetName;
	try
	{
		_Worksheet sheetTemp;
		sheetTemp.AttachDispatch(m_sheets.GetItem(COleVariant((long)(nIndex + 1))), true);
		strSheetName = sheetTemp.GetName();
		sheetTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return _T("");
	}
	return strSheetName;
}

int CPhdExcel::GetSheetIndexByName(LPCTSTR szSheetName)
{
	try
	{
		int nSheetCount = GetSheetCount();
		for (int i = 0; i < nSheetCount; i++)
		{
			CString strName = GetSheetNameByIndex(i);
			if (_tcscmp(szSheetName, strName) == 0)
			{
				return i;
			}
		}
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return -1;
	}
	
	return -1;
}

bool CPhdExcel::ModifyCurSheetName(LPCTSTR szNewSheetName)
{
	try
	{
		m_sheet.SetName(szNewSheetName);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::ModifySheetName(int nSheetIndex, LPCTSTR szNewSheetName)
{
	try
	{
		_Worksheet sheetTemp;
		sheetTemp.AttachDispatch(m_sheets.GetItem(COleVariant((long)(nSheetIndex + 1))), true);
		sheetTemp.SetName(szNewSheetName);
		sheetTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

int CPhdExcel::GetCurSheetIndex()
{
	CString strSheetName = m_sheet.GetName();
	int nIndex = GetSheetIndexByName(strSheetName);
	return nIndex;
}

bool CPhdExcel::AddSheet(LPCTSTR szSheetName)
{
	try
	{
		_Worksheet sheetTemp;
		sheetTemp.AttachDispatch(m_sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing),true);
		sheetTemp.SetName(szSheetName);
		sheetTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::DeleteSheet(int nSheetIndex)
{
	int nSheetCount = GetSheetCount();
	if (nSheetCount <= 1)
		return false;

	try
	{
		//�õ���ǰҳ������
		int nCurIndex = GetCurSheetIndex();
		if (nCurIndex == nSheetIndex)
		{
			m_sheet.Delete();
			m_sheet.ReleaseDispatch();
			//m_sheet�󶨵�һ��������
			m_sheet.AttachDispatch(m_sheets.GetItem(COleVariant((long)(1))), true);
			if (m_range)
				m_range.ReleaseDispatch();
			m_range.AttachDispatch(m_sheet.GetCells(), true);
		}
		else
		{
			_Worksheet sheetTemp;
			sheetTemp.AttachDispatch(m_sheets.GetItem(COleVariant((long)(nSheetIndex + 1))), true);
			sheetTemp.Delete();
			sheetTemp.ReleaseDispatch();
		}
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::DeleteSheet(LPCTSTR szSheetName)
{
	int nSheetCount = GetSheetCount();
	if (nSheetCount <= 1)
		return false;

	try
	{
		//�õ���ǰҳ������
		int nCurIndex = GetCurSheetIndex();
		int nSheetIndex = GetSheetIndexByName(szSheetName);
		if (nCurIndex == nSheetIndex)
		{
			m_sheet.Delete();
			m_sheet.ReleaseDispatch();
			//m_sheet�󶨵�һ��������
			m_sheet.AttachDispatch(m_sheets.GetItem(COleVariant((long)(1))), true);
			if (m_range)
				m_range.ReleaseDispatch();
			m_range.AttachDispatch(m_sheet.GetCells(), true);
		}
		else
		{
			_Worksheet sheetTemp;
			sheetTemp.AttachDispatch(m_sheets.GetItem(COleVariant((long)(nSheetIndex + 1))), true);
			sheetTemp.Delete();
			sheetTemp.ReleaseDispatch();
		}
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SwitchSheet(int nSheetIndex)
{
	try
	{
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		m_sheet.AttachDispatch(m_sheets.GetItem(COleVariant((long)(nSheetIndex + 1))), true);
		//
		if (m_range)
			m_range.ReleaseDispatch();
		m_range.AttachDispatch(m_sheet.GetCells(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SwitchSheet(LPCTSTR szSheetName)
{
	try
	{
		if (m_sheet)
			m_sheet.ReleaseDispatch();
		m_sheet.AttachDispatch(m_sheets.GetItem(_variant_t(szSheetName)), true);
		if (m_range)
			m_range.ReleaseDispatch();
		m_range.AttachDispatch(m_sheet.GetCells(), true);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetCurSheetActivate()
{
	try
	{
		m_sheet.Activate();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

int CPhdExcel::GetUsedRowCount()
{
	int nRowNum = 0;
	try
	{
		_ExRange rangeUsed;
		rangeUsed.AttachDispatch(m_sheet.GetUsedRange(),true);
		_ExRange rangeRow;
		rangeRow.AttachDispatch(rangeUsed.GetRows());
		nRowNum = rangeRow.GetCount();
		rangeUsed.ReleaseDispatch();
		rangeRow.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return 0;
	}
	return nRowNum;
}

int CPhdExcel::GetUsedColCount()
{
	int nColNum = 0;
	try
	{
		_ExRange rangeUsed;
		rangeUsed.AttachDispatch(m_sheet.GetUsedRange(), true);
		_ExRange rangeCol;
		rangeCol.AttachDispatch(rangeUsed.GetColumns());
		nColNum = rangeCol.GetCount();
		rangeUsed.ReleaseDispatch();
		rangeCol.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return 0;
	}
	return nColNum;
}

bool CPhdExcel::GetPrintArea(int& nRow1, int& nCol1, int& nRow2, int& nCol2)
{
	try
	{
		PageSetup pagesetupTemp;
		pagesetupTemp.AttachDispatch(m_sheet.GetPageSetup());
		CString strPrintArea = pagesetupTemp.GetPrintArea();
		pagesetupTemp.ReleaseDispatch();
		//
		int nIndex = strPrintArea.Find(_T(':'));
		if (nIndex == -1)
			return false;
		CString strLeft = strPrintArea.Left(nIndex);
		CString strRight = strPrintArea.Right(strPrintArea.GetLength() - nIndex - 1);
		//
		nIndex = strLeft.ReverseFind(_T('$'));
		if (nIndex == -1)
			return false;
		CString strRow1 = strLeft.Right(strLeft.GetLength() - nIndex - 1);
		strLeft = strLeft.Left(nIndex);
		nIndex = strLeft.ReverseFind(_T('$'));
		if (nIndex == -1)
			return false;
		CString strCol1 = strLeft.Right(strLeft.GetLength() - nIndex - 1);
		//
		nIndex = strRight.ReverseFind(_T('$'));
		if (nIndex == -1)
			return false;
		CString strRow2 = strRight.Right(strRight.GetLength() - nIndex - 1);
		strRight = strRight.Left(nIndex);
		nIndex = strRight.ReverseFind(_T('$'));
		if (nIndex == -1)
			return false;
		CString strCol2 = strRight.Right(strRight.GetLength() - nIndex - 1);
		//
		nRow1 = _wtoi(strRow1);
		nCol1 = UpperStrToNumber(strCol1);
		nRow2 = _wtoi(strRow2);
		nCol2 = UpperStrToNumber(strCol2);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

double CPhdExcel::GetRowHeight(int nRow, int nCol)
{
	double dHeight = 0;
	try
	{
		CString strCell = GetCellStr(nRow, nCol);
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell)));
		dHeight = rangeTemp.GetRowHeight().dblVal;
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return 0;
	}
	return dHeight;
}

double CPhdExcel::GetColWidth(int nRow, int nCol)
{
	double dWidth = 0;
	try
	{
		CString strCell = GetCellStr(nRow, nCol);
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell)));
		dWidth = rangeTemp.GetColumnWidth().dblVal;
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return 0;
	}
	return dWidth;
}

bool CPhdExcel::SetRowHeight(int nRow, double dHeight)
{
	try
	{
		CString strCell = GetCellStr(nRow, 1);
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell)));
		rangeTemp.SetRowHeight((_variant_t)dHeight);
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetColWidth(int nCol, double dWidth)
{
	try
	{
		CString strCell = GetCellStr(1, nCol);
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell)));
		rangeTemp.SetColumnWidth((_variant_t)dWidth);
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetAllRowHeight(double dHeight)
{
	try
	{
		m_range.SetRowHeight((_variant_t)dHeight);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetAllColWidth(double dWidth)
{
	try
	{
		m_range.SetColumnWidth((_variant_t)dWidth);
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

CString CPhdExcel::GetCellText(int nRow, int nCol)
{
	CString strText;
	SYSTEMTIME st;
	CString stry, strm, strd;

	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		_variant_t vtVal = rangeTemp.GetValue2();
		vtVal.ChangeType(VT_BSTR, NULL);
		switch (vtVal.vt)
		{
		case VT_BSTR:    //OLE Automation string
		{
			strText = vtVal.bstrVal;
			break;
		}
		case VT_R8: // 8 byte real
		{
			strText.Format(_T("%.f"), vtVal.dblVal);
			break;
		}
		case VT_DATE: //date
		{
			VariantTimeToSystemTime(vtVal.date, &st);
			stry.Format(_T("%d"), st.wYear);
			strm.Format(_T("%d"), st.wMonth);
			strd.Format(_T("%d"), st.wDay);
			strText = stry + _T("-") + strm + _T("-") + strd;
			break;
		}
		case VT_EMPTY: //empty
		{
			strText.Empty();
			break;
		}
		default:
		{
			strText.Empty();
			break;
		}
		}
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return _T("");
	}
	return strText;
}

bool CPhdExcel::SetCellText(int nRow, int nCol, LPCTSTR szText)
{
	try
	{
		m_range.SetItem(_variant_t(nRow), _variant_t(nCol), _variant_t(szText));
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::ClearContents()
{
	try
	{
		m_range.ClearContents();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetCellActivate(int nRow, int nCol)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		lpDisp = rangeTemp.Select();
		rangeTemp.Activate();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::MergeCell(int nCellRow1, int nCellCol1, int nCellRow2,
	int nCellCol2, bool bCenterAlign)
{
	try
	{
		CString strCell1 = GetCellStr(nCellRow1, nCellCol1);
		CString strCell2 = GetCellStr(nCellRow2, nCellCol2);
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(m_sheet.GetRange(_variant_t(strCell1),
			_variant_t(strCell2)));
		rangeTemp.Merge(_variant_t((long)0));
		if (bCenterAlign)
			rangeTemp.SetHorizontalAlignment(_variant_t((long)-4108));
		else
			rangeTemp.SetHorizontalAlignment(COleVariant((short)3));
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetAllCellCenterAlign()
{
	try
	{
		m_range.SetHorizontalAlignment(_variant_t((long)-4108));
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetAutoWrapText(int nRow, int nCol, bool bWrapText)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		rangeTemp.SetWrapText((_variant_t)(short)bWrapText);
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetFrame(int nCellRow1, int nCellCol1, int nCellRow2, int nCellCol2, bool bOuterFrame, bool bInnerFrame)
{
	try
	{
		CString strCell1 = GetCellStr(nCellRow1, nCellCol1);
		CString strCell2 = GetCellStr(nCellRow2, nCellCol2);
		_ExRange rangeTemp = m_sheet.GetRange(_variant_t(strCell1), _variant_t(strCell2));
		Borders borders;
		borders.AttachDispatch(rangeTemp.GetBorders());
		if (bOuterFrame)
		{
			Border bLeft;
			bLeft.AttachDispatch(borders.GetItem(7));
			bLeft.SetLineStyle(COleVariant((long)1));
			bLeft.SetWeight(COleVariant((long)3));

			Border bTop;
			bTop.AttachDispatch(borders.GetItem(8));
			bTop.SetLineStyle(COleVariant((long)1));
			bTop.SetWeight(COleVariant((long)3));

			Border bBottom;
			bBottom.AttachDispatch(borders.GetItem(9));
			bBottom.SetLineStyle(COleVariant((long)1));
			bBottom.SetWeight(COleVariant((long)3));

			Border bRight;
			bRight.AttachDispatch(borders.GetItem(10));
			bRight.SetLineStyle(COleVariant((long)1));
			bRight.SetWeight(COleVariant((long)3));

			bLeft.ReleaseDispatch();
			bTop.ReleaseDispatch();
			bBottom.ReleaseDispatch();
			bRight.ReleaseDispatch();
		}

		if (bInnerFrame)
		{
			Border bVertical;
			bVertical.AttachDispatch(borders.GetItem(11));
			bVertical.SetLineStyle(COleVariant((long)1));
			bVertical.SetWeight(COleVariant((long)2));

			Border bHorizontal;
			bHorizontal.AttachDispatch(borders.GetItem(12));
			bHorizontal.SetLineStyle(COleVariant((long)1));
			bHorizontal.SetWeight(COleVariant((long)2));

			bVertical.ReleaseDispatch();
			bHorizontal.ReleaseDispatch();
		}

		borders.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetFontSize(int nRow, int nCol, int nSize)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		_ExFont fontTemp;
		fontTemp.AttachDispatch(rangeTemp.GetFont());
		fontTemp.SetSize(COleVariant((long)nSize));//���������С
		fontTemp.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetBoldFont(int nRow, int nCol, bool bBold)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		_ExFont fontTemp;
		fontTemp.AttachDispatch(rangeTemp.GetFont());
		fontTemp.SetBold(COleVariant((long)(bBold ? 1 : 0)));//��������Ӵ�
		fontTemp.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetFontType(int nRow, int nCol, LPCTSTR szTextType)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(nCol));
		_ExRange rangeTemp;
		rangeTemp.AttachDispatch(lpDisp.pdispVal);
		_ExFont fontTemp;
		fontTemp.AttachDispatch(rangeTemp.GetFont());
		fontTemp.SetName(COleVariant(szTextType));
		fontTemp.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetCellTextColor(int nRow, int nCol, int nColorIndex)
{
	try
	{
		CString strCell = GetCellStr(nRow, nCol);
		_ExRange rangeTemp = m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell));
		_ExFont fontTemp;
		fontTemp.AttachDispatch(rangeTemp.GetFont());
		fontTemp.SetColorIndex(_variant_t(nColorIndex));
		fontTemp.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::SetCellBackgroundColor(int nRow, int nCol, int nColorIndex)
{
	try
	{
		CString strCell = GetCellStr(nRow, nCol);
		_ExRange rangeTemp = m_sheet.GetRange(_variant_t(strCell), _variant_t(strCell));
		Interior interTemp;
		interTemp.AttachDispatch(rangeTemp.GetInterior());
		interTemp.SetColorIndex(_variant_t(nColorIndex));
		interTemp.ReleaseDispatch();
		rangeTemp.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::InsertRow(int nRow)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(1));
		_ExRange copyFrom, copyTo;
		copyTo.AttachDispatch(lpDisp.pdispVal);
		copyFrom.AttachDispatch(copyTo.GetEntireRow());
		copyFrom.Insert(vtMissing);

		copyFrom.ReleaseDispatch();
		copyTo.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::InsertCol(int nCol)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(1), _variant_t(nCol));
		_ExRange copyFrom, copyTo;
		copyTo.AttachDispatch(lpDisp.pdispVal);
		copyFrom.AttachDispatch(copyTo.GetEntireColumn());
		copyFrom.Insert(vtMissing);

		copyFrom.ReleaseDispatch();
		copyTo.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::DeleteRow(int nRow)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(nRow), _variant_t(1));
		_ExRange copyFrom, copyTo;
		copyTo.AttachDispatch(lpDisp.pdispVal);
		copyFrom.AttachDispatch(copyTo.GetEntireRow());
		copyFrom.Delete(vtMissing);

		copyFrom.ReleaseDispatch();
		copyTo.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::DeleteCol(int nCol)
{
	try
	{
		VARIANT lpDisp = m_range.GetItem(_variant_t(1), _variant_t(nCol));
		_ExRange copyFrom, copyTo;
		copyTo.AttachDispatch(lpDisp.pdispVal);
		copyFrom.AttachDispatch(copyTo.GetEntireColumn());
		copyFrom.Delete(vtMissing);

		copyFrom.ReleaseDispatch();
		copyTo.ReleaseDispatch();
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

bool CPhdExcel::CloseCurOpenExcel(LPCTSTR szExcelPath)
{
	LPDISPATCH lpDisp = NULL;
	CoInitialize(NULL);
	CLSID clsid;
	HRESULT hr;
	hr = ::CLSIDFromProgID(_T("Excel.Application"), &clsid); //ͨ��ProgIDȡ��CLSID
	if (FAILED(hr))
	{
		return FALSE;//������û��excel����
	}

	IUnknown *pUnknown = NULL;
	IDispatch *pDispatch = NULL;
	hr = ::GetActiveObject(clsid, NULL, &pUnknown); //�����Ƿ��г���������
	if (FAILED(hr))
	{//û��excel����������
		return false;
	}
	else
	{//��excel����������
		try
		{
			_ExApplication appTemp;
			hr = pUnknown->QueryInterface(IID_IDispatch, (LPVOID *)&appTemp);
			if (FAILED(hr))
				throw(_T("û��ȡ��IDispatchPtr"));
			pUnknown->Release();
			pUnknown = NULL;

			lpDisp = appTemp.GetWorkbooks();
			Workbooks booksTemp;
			booksTemp.AttachDispatch(lpDisp, TRUE);
			int nLen = booksTemp.GetCount();
			CString strName;
			bool bIsFind = false;
			_Workbook bookTemp;
			for (int i = 1; i <= nLen; i++)
			{
				if (bookTemp)
					bookTemp.ReleaseDispatch();
				lpDisp = booksTemp.GetItem(_variant_t(i));
				bookTemp.AttachDispatch(lpDisp, TRUE);
				strName = bookTemp.GetFullName();
				
				if (_tcscmp(strName, szExcelPath) == 0)
				{
					bIsFind = true;
					break;
				}
			}
			if (!bIsFind)
			{
				bookTemp.ReleaseDispatch();
				booksTemp.ReleaseDispatch();
				appTemp.ReleaseDispatch();
				CoUninitialize();//�رյ�ǰ�߳��ϵ�com�⣬ж�ظ��̼߳��ص�����dll,�ͷŸ��߳�ά��������������Դ����ǿ�ƹرո��߳��ϵ�����RPC���ӡ�
				return false;//û�ҵ�
			}
				
			//���浱ǰ������
			appTemp.SetAlertBeforeOverwriting(FALSE);
			appTemp.SetDisplayAlerts(FALSE);
			bookTemp.Save();
			//�ر�ָ���Ĺ�����
			COleVariant covFalse((short)FALSE);
			COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			bookTemp.Close(covFalse, _variant_t(szExcelPath), covOptional);
			bookTemp.ReleaseDispatch();
			//booksTemp.Close();
			booksTemp.ReleaseDispatch();
			//appTemp.Quit();
			appTemp.ReleaseDispatch();
			CoUninitialize();//�رյ�ǰ�߳��ϵ�com�⣬ж�ظ��̼߳��ص�����dll,�ͷŸ��߳�ά��������������Դ����ǿ�ƹرո��߳��ϵ�����RPC���ӡ�
			return true;
		}
		catch (CException* e)
		{
			return FALSE;
		}
	}
}

bool CPhdExcel::ShowExcel()
{
	try
	{
		m_app.SetVisible(true);
		m_app.SetWindowState(SW_SHOWNORMAL);
		::BringWindowToTop(HWND(m_app.GetHwnd()));//���ô��ڶ�����ʾ
	}
	catch (CException* e)
	{
		TCHAR szError[1024];
		e->GetErrorMessage(szError, 1024);
		return false;
	}
	return true;
}

void CPhdExcel::NumberToUpperStr(int nNumber, CString& strNew)
{
	if (nNumber < 1)
		return;
	nNumber--;

	int num2 = nNumber / (25 + 1);
	int num3 = nNumber % (25 + 1);
	TCHAR newStr1 = TCHAR(num3 + 65);//��λ�ַ�
	strNew.Insert(0, newStr1);
	if (num2 == 0)
	{
		return;
	}
	else if (0 < num2 && num2 <= 25)
	{
		newStr1 = TCHAR(num2 + 64);//ǰλ�ַ�
		strNew.Insert(0, newStr1);
	}
	else if (num2 > 25)
	{
		NumberToUpperStr(num2, strNew);
	}
}

int CPhdExcel::UpperStrToNumber(LPCTSTR szStr)
{
	int nNumber = 0;
	CString str = szStr;
	int nCount = str.GetLength();
	int j = nCount;
	for (int i = 0; i < nCount; i++)
	{
		TCHAR curStr = str[j-1];
		if (curStr < _T('A') && curStr > _T('Z'))
			return 0;
		if (i == 0)
		{
			int nNumTemp = curStr;
			nNumber = nNumTemp - 64;
		}
		else
		{
			//��ָ��
			int nPow = std::pow(26.0, i);
			//
			int nNumTemp = curStr;
			int nNumTemp2 = (nNumTemp - 64) * nPow;
			
			nNumber += nNumTemp2;
		}
		j--;
	}

	return nNumber;
}

CString CPhdExcel::GetCellStr(int nRow, int nCol)
{
	CString strCol;
	NumberToUpperStr(nCol,strCol);
	CString strCell;
	strCell.Format(_T("%s%d"), strCol,nRow);
	return strCell;
}

bool CPhdExcel::IsNumber(LPCTSTR sz, bool bAcceptDouble)
{
	TCHAR cCurrent = 0;
	int   iDot = 0;
	CString sValue = sz;

	// check
	if (sValue.IsEmpty()) 
		return false;
	
	// Check each character of the string
	for (int i = 0; i < sValue.GetLength(); i++) 
	{
		cCurrent = sValue.GetAt(i);

		// The minus sign may only be at the begin
		if (cCurrent == _T('-') && i == 0) 
			continue;
		
		// A dot may not occure of we do not accept doubles
		if (cCurrent == _T('.') && !bAcceptDouble) 
			return false;
		
		// A dot may only occure once
		if (cCurrent == _T('.'))
		{
			iDot++;
			if (iDot == 1) 
				continue;
			else 
				return false;
		}

		// A number is something between 0 and 9, dah!
		if (cCurrent < _T('0') || cCurrent > _T('9')) 
			return false;
	}

	// We passed our wonderfull check,
	// so the string is an int or double Yiihaa
	return true;
}
