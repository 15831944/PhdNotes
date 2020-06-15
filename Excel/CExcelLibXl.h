#pragma once
#include "libxl.h"
using namespace libxl;

#define OK _T("ok")

class CExcelLibXl
{
public:
	CExcelLibXl();
	~CExcelLibXl();

#pragma region 初始化

	// Summary:   创建一个Excel文件
	// Time:	  2019年12月19日 peihaodong
	// Explain:	  
	bool CreateExcelFile(LPCTSTR szExcelPath);

	// Summary:   打开Excel
	// Time:	  2019年12月19日 peihaodong
	// Explain:	  
	bool OpenExcel(LPCTSTR szExcelPath);

	// Summary:   打开默认sheet页
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool OpenActiveSheet();

	bool Save();

	// Summary:   启动这个Excel文件
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool LaunchExcel() const;

#pragma endregion

#pragma region 读取数据

	// Summary:   得到打印区域
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool GetPrintArea(int* rowFirstIndex, int* rowLastIndex,
		int* colFirstIndex, int* colLastIndex) const;

	// Summary:   得到指定单元格的文本字符串
	// Time:	  2019年12月21日 peihaodong
	// Explain:	  nprec-如果该文本是浮点数，设置转为字符串的小数点位数
	CString GetText(int nRowIndex, int nColIndex,int nprec = 2) const;
	// Summary:   得到指定单元格的文本字符串
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	CString GetTextString(int nRowIndex,int nColIndex) const;
	// Summary:   得到指定单元格的文本数字
	// Time:	  2019年12月21日 peihaodong
	// Explain:	  
	double GetTextNumber(int nRowIndex, int nColIndex) const;
	// Summary:   得到指定单元格的日期
	// Time:	  2019年12月23日 peihaodong
	// Explain:	  
	CString GetTextTime(int nRowIndex, int nColIndex) const;

	// Summary:   得到已经被使用的行列
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  （返回的不是下标）
	int GetFirstRow() const;
	int GetLastRow() const;
	int GetFirstCol() const;
	int GetLastCol() const;

	// Summary:   得到字体指针
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  该字体存在，得到该字体的指针；不存在，创建该字体
	libxl::Font* GetFont(LPCTSTR szName) const;

	// Summary:   得到指定单元格的格式指针
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	libxl::Format* GetFormat(int nRowIndex,int nColIndex) const;

#pragma endregion

#pragma region 写入数据

	// Summary:   添加格式指针
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  在工作簿中创建一个新的格式
	libxl::Format* AddFormat() const;

	// Summary:   设置格式居中对齐
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	void SetFormatAlignCenter(libxl::Format* pFormat) const;

	// Summary:   设置指定单元格的文本字符串
	// Time:	  2019年12月21日 peihaodong
	// Explain:	  
	bool SetText(int nRowIndex, int nColIndex, LPCTSTR szText,
		libxl::Format* pFormat = NULL) const;
	// Summary:   设置指定单元格的文本字符串
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool SetTextString(int nRowIndex, int nColIndex, LPCTSTR szText,
		libxl::Format* pFormat = NULL) const;
	// Summary:   设置指定单元格的文本数字
	// Time:	  2019年12月21日 peihaodong
	// Explain:	  
	bool SetTextNumber(int nRowIndex, int nColIndex, double dNumber,
		libxl::Format* pFormat = NULL) const;
	// Summary:   设置指定单元格的日期
	// Time:	  2019年12月23日 peihaodong
	// Explain:	  
	bool SetTextTime(int nRowIndex, int nColIndex, 
		int nYear,int nMonth,int nDay,
		int nHour = 0,int nMin = 0,int nSec = 0,int nMsec = 0,
		libxl::Format* pFormat = NULL) const;

	// Summary:   设置列宽
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool SetColWidth(int colFirst, int colLast, double width,
		libxl::Format* pFormat = NULL) const;

	// Summary:   设置行高
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	bool SetRowHeight(int nRowIndex,double dHeight,
		libxl::Format* pFormat = NULL) const;

	// Summary:   设置列的格式
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool SetColFormat(int colFirst, int colLast,
		libxl::Format* pFormat) const;

	// Summary:   设置行的格式
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	bool SetRowFormat(int rowIndex,libxl::Format* pFormat) const;

	// Summary:   合并单元格
	// Time:	  2019年12月20日 peihaodong
	// Explain:	  
	bool MergeCells(int rowFirst, int rowLast, int colFirst,
		int colLast) const;

	// Summary:   在nRowIndex插入一行，原nRowIndex行会变为nRowIndex+1
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	bool InsertRow(int nRowIndex) const;
	// Summary:   在nRowFirstIndex~nRowLastIndex位置插入新的行
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	bool InsertRow(int nRowFirstIndex,int nRowLastIndex) const;

	// Summary:   设置打印区域
	// Time:	  2019年12月25日 peihaodong
	// Explain:	  
	bool SetPrintArea(int rowFirstIndex, int rowLastIndex,
		int colFirstIndex, int colLastIndex) const;

	// Summary:   清空该页
	// Time:	  2020年2月13日 peihaodong
	// Explain:	  
	void ClearSheetContents() const;

#pragma region 

private:
	CString GetSuffix(LPCTSTR szExcelPath) const;

	// Summary:   判断字串符是否是由数字组成
	// Time:	  2019年12月21日 peihaodong
	// Explain:	  
	bool StringIsNumber(LPCTSTR szString) const;

private:
	Book* m_pBook;				//工作簿(一个Excel文件）
	Sheet* m_pSheet;			//工作表(Excel文件中的一页)
// 	libxl::Format* m_pFormat;	//格式（格式包含字体）
// 	libxl::Font* m_pFont;		//字体

	CString m_strExcelPath;	
};

