#include "StdAfx.h"
#include "CPhdRegedit.h"


CPhdRegedit::CPhdRegedit(HKEY hKeyParent, LPCTSTR lpszKeyName)
	:m_hKeyParent(hKeyParent), m_lpszKeyName(lpszKeyName)
{
}


CPhdRegedit::~CPhdRegedit()
{
}

bool CPhdRegedit::SetValue(LPCTSTR lpszValueName, LPCTSTR lpszValue) const
{
	CRegKey reg;
	if (reg.Create(m_hKeyParent, m_lpszKeyName) != ERROR_SUCCESS)
		return false;

	if (reg.SetValue(lpszValue, lpszValueName) != ERROR_SUCCESS)
		return false;

	reg.Close(); 
	return true;
}

bool CPhdRegedit::SetValue(LPCTSTR lpszValueName, DWORD dwValue) const
{
	CRegKey reg;
	if (reg.Create(m_hKeyParent, m_lpszKeyName) != ERROR_SUCCESS)
		return false;

	if (reg.SetValue(dwValue, lpszValueName) != ERROR_SUCCESS)
		return false;

	reg.Close();
	return true;
}

bool CPhdRegedit::GetValue(LPCTSTR lpszValueName, CString& strValue) const
{
	CRegKey reg;
	if (reg.Open(m_hKeyParent, m_lpszKeyName) != ERROR_SUCCESS)
		return false;

	DWORD dwBufLen = 4096;
	CString str;
	if (reg.QueryValue(str.GetBuffer(4096), lpszValueName, &dwBufLen) != ERROR_SUCCESS)
		return false;
	str.ReleaseBuffer();
	strValue = str;

	return true;
}

bool CPhdRegedit::GetValue(LPCTSTR lpszValueName, DWORD& dwValue) const
{
	CRegKey reg;
	if (reg.Open(m_hKeyParent, m_lpszKeyName) != ERROR_SUCCESS)
		return false;

	if (reg.QueryValue(dwValue, lpszValueName) != ERROR_SUCCESS)
		return false;

	return true;
}

bool CPhdRegedit::DeleteChildKeyName(LPCTSTR lpszChildKeyName) const
{
	CRegKey reg;
	if (reg.Open(m_hKeyParent, m_lpszKeyName, KEY_ALL_ACCESS) != ERROR_SUCCESS)
		return false;

	if (reg.RecurseDeleteKey(lpszChildKeyName) != ERROR_SUCCESS)
		return false;

	reg.Close();
	return true;
}

bool CPhdRegedit::DeleteValue(LPCTSTR lpszValueName) const
{
	CRegKey reg;
	if (reg.Open(m_hKeyParent, m_lpszKeyName, KEY_ALL_ACCESS) != ERROR_SUCCESS)
		return false;

	if (reg.DeleteValue(lpszValueName) != ERROR_SUCCESS)
		return false;

	reg.Close();
	return true;
}

bool CPhdRegedit::GetAllChildKeyName(std::vector<CString>& vecStrName) const
{
	HKEY  hKeyResult = NULL;
	if (RegOpenKeyEx(m_hKeyParent, m_lpszKeyName, 0, KEY_READ | KEY_WOW64_64KEY, &hKeyResult) == ERROR_SUCCESS)
	{
		DWORD dwSubKeyCnt = 0;         // �Ӽ�������  
		DWORD dwSubKeyNameMaxLen = 0;  // �Ӽ����Ƶ���󳤶�(��������β��null�ַ�)  
		DWORD dwKeyValueCnt = 0;       // ��ֵ�������  
		DWORD dwKeyValueNameMaxLen = 0;// ��ֵ�����Ƶ���󳤶�(��������β��null�ַ�)  
		DWORD dwKeyValueDataMaxLen = 0;// ��ֵ�����ݵ���󳤶�(in bytes)  
		int ret = RegQueryInfoKey(hKeyResult, NULL, NULL, NULL, &dwSubKeyCnt, &dwSubKeyNameMaxLen, NULL, &dwKeyValueCnt, &dwKeyValueNameMaxLen,
			&dwKeyValueDataMaxLen, NULL, NULL);
		if (ret != ERROR_SUCCESS) // Error  
			return false;

		LPTSTR lpszSubKeyName = new TCHAR[dwSubKeyNameMaxLen + 1];
		for (DWORD index = 0; index < dwSubKeyCnt; ++index)
		{
			memset(lpszSubKeyName, 0, sizeof(TCHAR)*(dwSubKeyNameMaxLen + 1));
			DWORD dwNameCnt = dwSubKeyNameMaxLen + 1;
			int ret = RegEnumKeyEx(hKeyResult, index, lpszSubKeyName, &dwNameCnt, NULL, NULL, NULL, NULL);
			if (ret != ERROR_SUCCESS)
			{
				delete[] lpszSubKeyName;
				return false;
			}
			CString strName = lpszSubKeyName;
			vecStrName.push_back(strName);
		}
		delete[] lpszSubKeyName;
	}

	//�ر�ע���
	::RegCloseKey(hKeyResult);
	return true;
}
