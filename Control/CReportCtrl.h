#pragma once

struct sTitle
{
	CString _strTitle;
	std::vector<CString> _vCombo;
	int _nWidth;
	bool _bEditable;
	bool _bSortable;
	sTitle(CString strTitle)
	{
		_strTitle = strTitle;
		_bEditable = false;
		_nWidth = 3;
		_bSortable = false;
	}
};
typedef std::vector<sTitle>  XTP_TITLES;

class CReportRowData :public CXTPReportRecord
{
public:
	CReportRowData(const std::vector<CString> &vData)
	{
		SetEditable(TRUE);
		for (std::vector<CString>::const_iterator it = vData.begin(); it != vData.end(); it++)
		{
			CXTPReportRecordItem *pItem = AddItem(new CXTPReportRecordItemText(*it));
			pItem->SetEditable(TRUE);
		}
	}
};

const CString REPORT_TITLE[] = { _T("����"),_T("��Ϣ") };
const int REPORT_TITLE_W[] = { 1,2 };

class CReportCtrl :public CXTPReportControl
{
public:
	CReportCtrl();
	~CReportCtrl();

	void InitCtrol();//��ʼ���ؼ�
	bool SetDataText(int nRowIndex, int nColIndex, const CString& strText);//����ָ�����е��ı�
	int GetRowCount() const;//�õ�����

	//����һ��
	void AppendRow(const CString& strText);
	void AppendRow(const std::vector<CString> &vData);
	void AppendRow(const CString& strTag, const CString& strText);

	//�õ�ָ�����е��ı�
	CString GetData(int nRowIndex, int nColIndex);
	int GetRowData(int nIndex, std::vector<CString>& vData);
};

