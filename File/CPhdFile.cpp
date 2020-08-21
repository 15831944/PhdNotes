#include "StdAfx.h"
#include "CPhdFile.h"
#include <locale> 


CPhdFile::CPhdFile(const CString& strFilePath)
	:m_strFilePath(strFilePath)
{
}


CPhdFile::~CPhdFile()
{
}

bool CPhdFile::WriteString(const CString& str) const
{
	CStdioFile file;
	CFileException error;
	setlocale(LC_CTYPE, ("chs"));
	if (!file.Open(m_strFilePath, CFile::modeWrite | CFile::modeCreate | CFile::modeNoTruncate, &error))
		//CFile::modeCreate | CFile::modeNoTruncate的组合使添加文件时追加到已有文件末尾
		return false;

	file.SeekToEnd();

	CString strData = str;
	strData += _T("\n");//相当于回车字符
	file.WriteString(strData);

	file.Close();
	return TRUE;
}

bool CPhdFile::WriteString(const std::vector<CString>& vecstr) const
{
	CStdioFile file;
	CFileException error;
	setlocale(LC_CTYPE, ("chs"));
	if (!file.Open(m_strFilePath, CFile::modeWrite | CFile::modeCreate | CFile::modeNoTruncate, &error))
		//CFile::modeCreate | CFile::modeNoTruncate的组合使添加文件时追加到已有文件末尾
		return false;

	file.SeekToEnd();

	for (int i = 0; i < vecstr.size(); ++i)
	{
		CString strData = vecstr[i];
		strData += _T("\n");//相当于回车字符
		//strData += _T("\r\n");//相当于回车字符
		file.WriteString(strData);
	}

	file.Close();
	return TRUE;
}

bool CPhdFile::ReadString(std::vector<CString>& vecStr) const
{
	CStdioFile file;
	CFileException error;
	setlocale(LC_CTYPE, ("chs"));
	if (!file.Open(m_strFilePath, CFile::modeRead | CFile::shareDenyNone, &error))
		return false;

	CString strLine;
	while (file.ReadString(strLine))
	{
		strLine.Trim();//清除开头和结尾的空格
		if (!strLine.IsEmpty())
			vecStr.push_back(strLine);
	}

	file.Close();
	return true;
}
