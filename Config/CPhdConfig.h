#pragma once


//�����ļ�
class CPhdConfig
{
	/*�����ļ���Ҫ��Ϊ�����֣��ֶ���(section)	����(key)	��ֵ(value)*/
public:
	CPhdConfig();
	CPhdConfig(const CString& strFilePath);
	~CPhdConfig();

	//************************************
	// Summary: 	�����ļ�·��
	// Parameters: 	
	//    szFilePath		- 	
	// Returns:   	
	//************************************
	void SetFilePath(LPCTSTR szFilePath);
	
	//************************************
	// Summary: 	д���ֵ
	// Parameters: 	
	//    szSection		- 	�����ֶ���
	//    szName		- 	�������
	//    szValue		- 	�����ֵ
	// Returns:   	
	//************************************
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, LPCTSTR szValue) const;
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, int value) const;
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, double value) const;

	//************************************
	// Summary: 	��ȡ��ֵ
	// Parameters: 	
	//    szSection		- 	�����ֶ���
	//    szName		- 	�������
	//    szDefault		- 	����Ĭ��ֵ����������ļ���û�ж������򷵻�Ĭ��ֵ��
	// Returns:   	
	//************************************
	CString GetValue(LPCTSTR szSection, LPCTSTR szKey, LPCTSTR szDefault = nullptr) const;
	int GetValue(LPCTSTR szSection, LPCTSTR szKey, int nDefault) const;
	double GetValue(LPCTSTR szSection, LPCTSTR szKey, double fDefault) const;

	//************************************
	// Summary: 	�õ��������ļ��������ֶ���
	// Parameters: 	
	// Returns:   	
	//************************************
	std::vector<CString> GetAllSections() const;

	//************************************
	// Summary: 	ɾ��ĳһ�ֶ�
	// Parameters: 	
	//    lpSection		- 	�����ֶ���
	// Returns:   	
	//************************************
	BOOL DelSection(LPCTSTR lpSection) const;

	//************************************
	// Summary: 	ɾ�������ֶ�
	// Parameters: 	
	// Returns:   	
	//************************************
	BOOL DelAllSections() const;

	//************************************
	// Summary: 	ɾ������
	// Parameters: 	
	//    lpSection		- 	�����ֶ���
	//    lpKey		- 	�������
	// Returns:   	
	//************************************
	BOOL DelKey(LPCTSTR lpSection, LPCTSTR lpKey) const;

	//************************************
	// Summary: 	�õ��ֶ������еļ����ͼ�ֵ
	// Parameters: 	
	//    lpszSection		- 	�����ֶ���
	//    vecKey		- 	�����������
	//    vecValue		- 	�����ֵ����
	// Returns:   	�����ֶ��¼�������
	//************************************
	int GetSectionData(LPCTSTR lpszSection, std::vector<CString>& vecKey, std::vector<CString>& vecValue) const;

protected:
	CString m_strFilePath;
	CString m_strAppdataFile;	// ������appdataĿ¼��

	CString GetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR sz_default) const;
	bool SetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR value) const;

private:
	CString GetAppdataDir() const;
};


