#pragma once
#include "libxl.h"
using namespace libxl;

/***********************************************
   >   Class Name: PhdLibXLExcel
   >     Describe: 通过第三方库LibXL，来读写excel文件，该类是对LibXL接口的封装
   >       Author: peihaodong
   > Created Time: 2020年9月27日
   >         Blog: https://blog.csdn.net/phd17621680432
**********************************************/
class PhdLibXLExcel 
{
public:
	PhdLibXLExcel();
	~PhdLibXLExcel();

public:

	//************************************
	// Summary:  创建excel文件
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool CreateExcel(LPCTSTR szExcelPath);

	//************************************
	// Summary:  打开文件
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool OpenExcel(LPCTSTR szExcelPath);

	//************************************
	// Summary:  打开sheet页
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool OpenActiveSheet();
	bool OpenSheetByIndex(int nSheetIndex);
	bool OpenSheetByName(LPCTSTR szSheetName);

	//************************************
	// Summary:  保存excel
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool Save() const {
		return m_pBook->save(m_strExcelPath);
	}

	//************************************
	// Summary:  启动excel
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool LaunchExcel() const;

#pragma region 读取数据

	//************************************
	// Summary:  得到指定单元格文本和格式
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool GetCellTextOfStr(int nRowIndex, int nColIndex,CString& strText, libxl::Format** format = 0) const;
	bool GetCellTextOfNum(int nRowIndex, int nColIndex, double& dText, libxl::Format** format = 0) const;
	bool GetCellTextOfDate(int nRowIndex, int nColIndex, CString& strDate, libxl::Format** format = 0) const;

	//************************************
	// Summary:  得到打印区域
	// Explain:  
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool GetPrintArea(int* rowFirstIndex, int* rowLastIndex,
		int* colFirstIndex, int* colLastIndex) const;

	//************************************
	// Summary:  得到有效行数
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	int GetUsedRowCount() const;

	//************************************
	// Summary:  得到有效列数
	// Explain:     
	// Time:      2020年9月27日 peihaodong
	//************************************
	int GetUsedColCount() const;

	//************************************
	// Summary:  得到字体指针
	// Explain:  该字体存在，得到该字体的指针；不存在，创建该字体
	// Time:      2020年9月27日 peihaodong
	//************************************
	libxl::Font* GetFont(LPCTSTR szFontName) const;

	//************************************
	// Summary:  得到指定单元格的格式指针
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	libxl::Format* GetFormat(int nRowIndex, int nColIndex) const;
#pragma endregion

#pragma region 写入数据

	//************************************
	// Summary:  设置单元格文本和格式
	// Explain:   format = 0时忽略设置格式
	// Time:      2020年9月27日 peihaodong
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
	// Summary:  得到一个新的格式
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	libxl::Format* GetNewFormat() const;

	//************************************
	// Summary:  设置格式居中对齐
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	void SetFormatAlignCenter(libxl::Format* pFormat) const;

	//************************************
	// Summary:  设置指定行的行高和格式
	// Explain:   format = 0时忽略设置格式
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool SetRowHeight(int nRowIndex,double dHeight, libxl::Format* format = 0) const;

	//************************************
	// Summary:  设置指定列的列宽和格式
	// Explain:   format = 0时忽略设置格式
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool SetColWidth(int nColFirst,int nColLast,double dWidth, libxl::Format* format = 0) const;

	//************************************
	// Summary:  合并单元格
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool MergeCells(int nRowFirst, int nRowLast, int nColFirst,int nColLast) const;

	//************************************
	// Summary:  在nRowIndex插入一行，原nRowIndex行会变为nRowIndex+1
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool InsertRow(int nRowIndex) const;
	//************************************
	// Summary:  在nRowFirstIndex~nRowLastIndex位置插入新的行
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	bool InsertRow(int nRowFirstIndex, int nRowLastIndex) const;

	//************************************
	// Summary:  设置打印区域
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	void SetPrintArea(int rowFirstIndex, int rowLastIndex,
		int colFirstIndex, int colLastIndex) const;

	//************************************
	// Summary:  清空该页
	// Explain:   
	// Time:      2020年9月27日 peihaodong
	//************************************
	void ClearSheetContents() const;

#pragma endregion

protected:
	Book* m_pBook;				//工作簿(一个Excel文件）
	Sheet* m_pSheet;			//工作表(Excel文件中的一页)
// 	libxl::Format* m_pFormat;	//格式（格式包含字体）
// 	libxl::Font* m_pFont;		//字体

	CString m_strExcelPath;
};