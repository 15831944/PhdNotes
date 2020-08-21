#pragma once
#include <vector>


class CPhdFile
{
public:
	CPhdFile(const CString& strFilePath);
	~CPhdFile();

	//************************************
	// Summary: 	�ļ���д���ַ���
	// Parameters: 	
	//    str		- 	�����ַ���
	// Returns:   	
	//************************************
	bool WriteString(const CString& str) const;
	bool WriteString(const std::vector<CString>& vecstr) const;

	//************************************
	// Summary: 	��ȡ�ļ��е��ַ���
	// Parameters: 	
	//    vecStr		- 	����ַ�������
	// Returns:   	
	//************************************
	bool ReadString(std::vector<CString>& vecStr) const;

protected:
	CString m_strFilePath;
};

