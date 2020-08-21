#pragma once
#include <vector>


class CPhdFile
{
public:
	CPhdFile(const CString& strFilePath);
	~CPhdFile();

	//************************************
	// Summary: 	文件中写入字符串
	// Parameters: 	
	//    str		- 	输入字符串
	// Returns:   	
	//************************************
	bool WriteString(const CString& str) const;
	bool WriteString(const std::vector<CString>& vecstr) const;

	//************************************
	// Summary: 	读取文件中的字符串
	// Parameters: 	
	//    vecStr		- 	输出字符串集合
	// Returns:   	
	//************************************
	bool ReadString(std::vector<CString>& vecStr) const;

protected:
	CString m_strFilePath;
};

