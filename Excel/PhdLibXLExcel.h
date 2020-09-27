#pragma once
#include "libxl.h"
using namespace libxl;

/***********************************************
   >   Class Name: PhdLibXLExcel
   >     Describe: ͨ����������LibXL������дexcel�ļ��������Ƕ�LibXL�ӿڵķ�װ
   >       Author: peihaodong
   > Created Time: 2020��9��27��
   >         Blog: https://blog.csdn.net/phd17621680432
**********************************************/
class PhdLibXLExcel 
{
public:
	PhdLibXLExcel();
	~PhdLibXLExcel();

public:

	//************************************
	// Summary:  ����excel�ļ�
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool CreateExcel(LPCTSTR szExcelPath);

	//************************************
	// Summary:  ���ļ�
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool OpenExcel(LPCTSTR szExcelPath);

	//************************************
	// Summary:  ��sheetҳ
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool OpenActiveSheet();
	bool OpenSheetByIndex(int nSheetIndex);
	bool OpenSheetByName(LPCTSTR szSheetName);

	//************************************
	// Summary:  ����excel
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool Save() const {
		return m_pBook->save(m_strExcelPath);
	}

	//************************************
	// Summary:  ����excel
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool LaunchExcel() const;

#pragma region ��ȡ����

	//************************************
	// Summary:  �õ�ָ����Ԫ���ı��͸�ʽ
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool GetCellTextOfStr(int nRowIndex, int nColIndex,CString& strText, libxl::Format** format = 0) const;
	bool GetCellTextOfNum(int nRowIndex, int nColIndex, double& dText, libxl::Format** format = 0) const;
	bool GetCellTextOfDate(int nRowIndex, int nColIndex, CString& strDate, libxl::Format** format = 0) const;

	//************************************
	// Summary:  �õ���ӡ����
	// Explain:  
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool GetPrintArea(int* rowFirstIndex, int* rowLastIndex,
		int* colFirstIndex, int* colLastIndex) const;

	//************************************
	// Summary:  �õ���Ч����
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	int GetUsedRowCount() const;

	//************************************
	// Summary:  �õ���Ч����
	// Explain:     
	// Time:      2020��9��27�� peihaodong
	//************************************
	int GetUsedColCount() const;

	//************************************
	// Summary:  �õ�����ָ��
	// Explain:  ��������ڣ��õ��������ָ�룻�����ڣ�����������
	// Time:      2020��9��27�� peihaodong
	//************************************
	libxl::Font* GetFont(LPCTSTR szFontName) const;

	//************************************
	// Summary:  �õ�ָ����Ԫ��ĸ�ʽָ��
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	libxl::Format* GetFormat(int nRowIndex, int nColIndex) const;
#pragma endregion

#pragma region д������

	//************************************
	// Summary:  ���õ�Ԫ���ı��͸�ʽ
	// Explain:   format = 0ʱ�������ø�ʽ
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool SetCellTextOfStr(int nRowIndex, int nColIndex, LPCTSTR szText,
		libxl::Format* pFormat = 0) const;
	bool SetCellTextOfNum(int nRowIndex, int nColIndex, double dNumber,
		libxl::Format* pFormat = 0) const;
	bool SetCellTextOfDate(int nRowIndex, int nColIndex,
		int nYear, int nMonth, int nDay,
		int nHour = 0, int nMin = 0, int nSec = 0, int nMsec = 0,
		libxl::Format* pFormat = 0) const;

	//************************************
	// Summary:  �õ�һ���µĸ�ʽ
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	libxl::Format* GetNewFormat() const;

	//************************************
	// Summary:  ���ø�ʽ���ж���
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	void SetFormatAlignCenter(libxl::Format* pFormat) const;

	//************************************
	// Summary:  ����ָ���е��иߺ͸�ʽ
	// Explain:   format = 0ʱ�������ø�ʽ
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool SetRowHeight(int nRowIndex,double dHeight, libxl::Format* format = 0) const;

	//************************************
	// Summary:  ����ָ���е��п�͸�ʽ
	// Explain:   format = 0ʱ�������ø�ʽ
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool SetColWidth(int nColFirst,int nColLast,double dWidth, libxl::Format* format = 0) const;

	//************************************
	// Summary:  �ϲ���Ԫ��
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool MergeCells(int nRowFirst, int nRowLast, int nColFirst,int nColLast) const;

	//************************************
	// Summary:  ��nRowIndex����һ�У�ԭnRowIndex�л��ΪnRowIndex+1
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool InsertRow(int nRowIndex) const;
	//************************************
	// Summary:  ��nRowFirstIndex~nRowLastIndexλ�ò����µ���
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	bool InsertRow(int nRowFirstIndex, int nRowLastIndex) const;

	//************************************
	// Summary:  ���ô�ӡ����
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	void SetPrintArea(int rowFirstIndex, int rowLastIndex,
		int colFirstIndex, int colLastIndex) const;

	//************************************
	// Summary:  ��ո�ҳ
	// Explain:   
	// Time:      2020��9��27�� peihaodong
	//************************************
	void ClearSheetContents() const;

#pragma endregion

protected:
	Book* m_pBook;				//������(һ��Excel�ļ���
	Sheet* m_pSheet;			//������(Excel�ļ��е�һҳ)
// 	libxl::Format* m_pFormat;	//��ʽ����ʽ�������壩
// 	libxl::Font* m_pFont;		//����

	CString m_strExcelPath;
};