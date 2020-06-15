#include "StdAfx.h"
#include "CReportCtrl.h"


CReportCtrl::CReportCtrl()
{
}


CReportCtrl::~CReportCtrl()
{
}

void CReportCtrl::InitCtrol()
{
	AllowEdit(TRUE);
	XTP_TITLES zhTitle;

	for (int i = 0; i < sizeof(REPORT_TITLE) / sizeof(CString); i++)
	{
		sTitle tData(REPORT_TITLE[i]);
		tData._nWidth = REPORT_TITLE_W[i];
		zhTitle.push_back(tData);
	}

	int nIndex = 0;
	for (auto it = zhTitle.begin(); it != zhTitle.end(); it++)
	{
		auto pColumn = this->AddColumn(new CXTPReportColumn(nIndex++, it->_strTitle, it->_nWidth, XTP_REPORT_NOICON, NULL));
		pColumn->SetSortable(it->_bSortable);
		pColumn->SetEditable(TRUE);
		if (nIndex == 1)
		{
			pColumn->SetTreeColumn(TRUE);
			pColumn->SetEditable(FALSE);
		}
		if (it->_vCombo.size() > 0)
		{
			CXTPReportRecordItemEditOptions *pEditOptions = pColumn->GetEditOptions();
			for (auto itor = it->_vCombo.begin(); itor != it->_vCombo.end(); it++)
			{
				pEditOptions->AddConstraint(*itor);
			}
			pEditOptions->m_bAllowEdit = true;
			pEditOptions->m_bConstraintEdit = false;
			pEditOptions->AddComboButton(TRUE);
		}
	}
}

bool CReportCtrl::SetDataText(int nRowIndex, int nColIndex, const CString& strText)
{
	auto pRcd = GetRows()->GetAt(nRowIndex);
	if (pRcd == NULL)
		return false;
	auto pSetting = static_cast<CReportRowData*>(pRcd->GetRecord());
	if (pSetting == NULL)
		return false;
	int nNum = pSetting->GetItemCount();

	auto pItem = static_cast<CXTPReportRecordItemText*>(pSetting->GetItem(nColIndex));
	pItem->SetValue(strText);
	return true;
}

int CReportCtrl::GetRowCount() const
{
	return GetRows()->GetCount();
}

void CReportCtrl::AppendRow(const CString& strText)
{
	std::vector<CString> vInf;
	vInf.push_back(strText);
	int nCol = GetColumns()->GetCount();

	for (int i = vInf.size(); i < nCol; i++)
	{
		vInf.push_back(CString());
	}
	AddRecord(new CReportRowData(vInf));

	Populate();
}

void CReportCtrl::AppendRow(const std::vector<CString> &vData)
{
	std::vector<CString> vInf = vData;
	int nCol = GetColumns()->GetCount();//总共有几列

	for (int i = vInf.size(); i < nCol; i++)
	{
		vInf.push_back(CString());
	}
	AddRecord(new CReportRowData(vInf));

	Populate();
}

void CReportCtrl::AppendRow(const CString& strTag, const CString& strText)
{
	std::vector<CString> vecStr;
	vecStr.push_back(strTag);
	vecStr.push_back(strText);
	AppendRow(vecStr);
}

CString CReportCtrl::GetData(int nRowIndex, int nColIndex)
{
	CString strText;
	auto pRcd = GetRows()->GetAt(nRowIndex);
	if (pRcd == NULL)
		return strText;
	auto pSetting = static_cast<CReportRowData*>(pRcd->GetRecord());
	if (pSetting == NULL)
		return strText;
	int nNum = pSetting->GetItemCount();

	auto pItem = static_cast<CXTPReportRecordItemText*>(pSetting->GetItem(nColIndex));
	strText = pItem->GetValue();

	return strText;
}

int CReportCtrl::GetRowData(int nIndex, std::vector<CString>& vData)
{
	vData.clear();
	auto pRcd = GetRows()->GetAt(nIndex);
	if (pRcd == NULL)
	{
		return 0;
	}
	auto pSetting = static_cast<CXTPReportRecord*>(pRcd->GetRecord());
	if (pSetting == NULL)
	{
		return 0;
	}
	for (int i = 0; i < pSetting->GetItemCount(); i++)
	{
		auto pItem = static_cast<CXTPReportRecordItemText*>(pSetting->GetItem(i));
		vData.push_back(pItem->GetValue());
	}
	return vData.size();
}
