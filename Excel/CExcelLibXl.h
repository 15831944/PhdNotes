#pragma once
#include "libxl.h"
using namespace libxl;

#define OK _T("ok")

class CExcelLibXl
{
public:
	CExcelLibXl();
	~CExcelLibXl();

#pragma region ��ʼ��

	// Summary:   ����һ��Excel�ļ�
	// Time:	  2019��12��19�� peihaodong
	// Explain:	  
	bool CreateExcelFile(LPCTSTR szExcelPath);

	// Summary:   ��Excel
	// Time:	  2019��12��19�� peihaodong
	// Explain:	  
	bool OpenExcel(LPCTSTR szExcelPath);

	// Summary:   ��Ĭ��sheetҳ
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool OpenActiveSheet();

	bool Save();

	// Summary:   �������Excel�ļ�
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool LaunchExcel() const;

#pragma endregion

#pragma region ��ȡ����

	// Summary:   �õ���ӡ����
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool GetPrintArea(int* rowFirstIndex, int* rowLastIndex,
		int* colFirstIndex, int* colLastIndex) const;

	// Summary:   �õ�ָ����Ԫ����ı��ַ���
	// Time:	  2019��12��21�� peihaodong
	// Explain:	  nprec-������ı��Ǹ�����������תΪ�ַ�����С����λ��
	CString GetText(int nRowIndex, int nColIndex,int nprec = 2) const;
	// Summary:   �õ�ָ����Ԫ����ı��ַ���
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	CString GetTextString(int nRowIndex,int nColIndex) const;
	// Summary:   �õ�ָ����Ԫ����ı�����
	// Time:	  2019��12��21�� peihaodong
	// Explain:	  
	double GetTextNumber(int nRowIndex, int nColIndex) const;
	// Summary:   �õ�ָ����Ԫ�������
	// Time:	  2019��12��23�� peihaodong
	// Explain:	  
	CString GetTextTime(int nRowIndex, int nColIndex) const;

	// Summary:   �õ��Ѿ���ʹ�õ�����
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  �����صĲ����±꣩
	int GetFirstRow() const;
	int GetLastRow() const;
	int GetFirstCol() const;
	int GetLastCol() const;

	// Summary:   �õ�����ָ��
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  ��������ڣ��õ��������ָ�룻�����ڣ�����������
	libxl::Font* GetFont(LPCTSTR szName) const;

	// Summary:   �õ�ָ����Ԫ��ĸ�ʽָ��
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	libxl::Format* GetFormat(int nRowIndex,int nColIndex) const;

#pragma endregion

#pragma region д������

	// Summary:   ��Ӹ�ʽָ��
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  �ڹ������д���һ���µĸ�ʽ
	libxl::Format* AddFormat() const;

	// Summary:   ���ø�ʽ���ж���
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	void SetFormatAlignCenter(libxl::Format* pFormat) const;

	// Summary:   ����ָ����Ԫ����ı��ַ���
	// Time:	  2019��12��21�� peihaodong
	// Explain:	  
	bool SetText(int nRowIndex, int nColIndex, LPCTSTR szText,
		libxl::Format* pFormat = NULL) const;
	// Summary:   ����ָ����Ԫ����ı��ַ���
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool SetTextString(int nRowIndex, int nColIndex, LPCTSTR szText,
		libxl::Format* pFormat = NULL) const;
	// Summary:   ����ָ����Ԫ����ı�����
	// Time:	  2019��12��21�� peihaodong
	// Explain:	  
	bool SetTextNumber(int nRowIndex, int nColIndex, double dNumber,
		libxl::Format* pFormat = NULL) const;
	// Summary:   ����ָ����Ԫ�������
	// Time:	  2019��12��23�� peihaodong
	// Explain:	  
	bool SetTextTime(int nRowIndex, int nColIndex, 
		int nYear,int nMonth,int nDay,
		int nHour = 0,int nMin = 0,int nSec = 0,int nMsec = 0,
		libxl::Format* pFormat = NULL) const;

	// Summary:   �����п�
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool SetColWidth(int colFirst, int colLast, double width,
		libxl::Format* pFormat = NULL) const;

	// Summary:   �����и�
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	bool SetRowHeight(int nRowIndex,double dHeight,
		libxl::Format* pFormat = NULL) const;

	// Summary:   �����еĸ�ʽ
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool SetColFormat(int colFirst, int colLast,
		libxl::Format* pFormat) const;

	// Summary:   �����еĸ�ʽ
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	bool SetRowFormat(int rowIndex,libxl::Format* pFormat) const;

	// Summary:   �ϲ���Ԫ��
	// Time:	  2019��12��20�� peihaodong
	// Explain:	  
	bool MergeCells(int rowFirst, int rowLast, int colFirst,
		int colLast) const;

	// Summary:   ��nRowIndex����һ�У�ԭnRowIndex�л��ΪnRowIndex+1
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	bool InsertRow(int nRowIndex) const;
	// Summary:   ��nRowFirstIndex~nRowLastIndexλ�ò����µ���
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	bool InsertRow(int nRowFirstIndex,int nRowLastIndex) const;

	// Summary:   ���ô�ӡ����
	// Time:	  2019��12��25�� peihaodong
	// Explain:	  
	bool SetPrintArea(int rowFirstIndex, int rowLastIndex,
		int colFirstIndex, int colLastIndex) const;

	// Summary:   ��ո�ҳ
	// Time:	  2020��2��13�� peihaodong
	// Explain:	  
	void ClearSheetContents() const;

#pragma region 

private:
	CString GetSuffix(LPCTSTR szExcelPath) const;

	// Summary:   �ж��ִ����Ƿ������������
	// Time:	  2019��12��21�� peihaodong
	// Explain:	  
	bool StringIsNumber(LPCTSTR szString) const;

private:
	Book* m_pBook;				//������(һ��Excel�ļ���
	Sheet* m_pSheet;			//������(Excel�ļ��е�һҳ)
// 	libxl::Format* m_pFormat;	//��ʽ����ʽ�������壩
// 	libxl::Font* m_pFont;		//����

	CString m_strExcelPath;	
};

