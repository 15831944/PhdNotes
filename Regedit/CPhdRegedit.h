#pragma once

/*
���ӣ�
	CPhdRegedit reg(HKEY_CURRENT_USER, _T("CNPE"));
	reg.GetRegInt(_T("Balloon_Adjust"), m_bAdjust);
	reg.SetRegInt(_T("Balloon_Adjust"), m_bAdjust);
*/

//ע���
class CPhdRegedit
{
	//ע���  ��Ҫ��Ϊ����(hKeyParent)���Ӽ�(lpszKeyName)�ͼ�ֵ��������   
	//��ֵ���������������ɣ��ֱ�Ϊ������(lpszValueName)�����͡�����(lpszValue)
	//hKeyParent��	HKEY_CLASSES_ROOT		HKEY_CURRENT_USER		HKEY_LOCAL_MACHINE		HKEY_USERS
public:
	CPhdRegedit(HKEY hKeyParent,LPCTSTR lpszKeyName);
	~CPhdRegedit();

	//************************************
	// Summary: 	���ü�ֵ���ֵ
	// Parameters: 	
	//    lpszValueName		- 	�����ֵ�������
	//    lpszValue		- 	�����ֵ���ֵ
	// Returns:   	
	//************************************
	bool SetValue(LPCTSTR lpszValueName,LPCTSTR lpszValue) const;
	bool SetValue(LPCTSTR lpszValueName, DWORD dwValue) const;
	
	//************************************
	// Summary: 	�õ���ֵ���ֵ
	// Parameters: 	
	//    lpszValueName		- 	�����ֵ�������
	//    szValue		- 	�����ֵ���ֵ
	// Returns:   	
	//************************************
	bool GetValue(LPCTSTR lpszValueName, CString& strValue) const;
	bool GetValue(LPCTSTR lpszValueName, DWORD& dwValue) const;

	//************************************
	// Summary: 	ɾ���Ӽ���
	// Parameters: 	
	//    lpszKeyName		- 	�����Ӽ����Ӽ������ƣ�ֱ���������֣�������·����
	// Returns:   	
	//************************************
	bool DeleteChildKeyName(LPCTSTR lpszChildKeyName) const;

	//************************************
	// Summary: 	ɾ����ֵ��
	// Parameters: 	
	//    lpszValueName		- 	�����ֵ�������
	// Returns:   	
	//************************************
	bool DeleteValue(LPCTSTR lpszValueName) const;

	//************************************
	// Summary: 	�õ��Ӽ��������е�������
	// Parameters: 	
	//    vecStrName		- 	�������������
	// Returns:   	
	//************************************
	bool GetAllChildKeyName(std::vector<CString>& vecStrName) const;

protected:
	HKEY m_hKeyParent; 
	LPCTSTR m_lpszKeyName;
	
};

