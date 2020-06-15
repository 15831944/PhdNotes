#pragma once

/*
例子：
	CPhdRegedit reg(HKEY_CURRENT_USER, _T("CNPE"));
	reg.GetRegInt(_T("Balloon_Adjust"), m_bAdjust);
	reg.SetRegInt(_T("Balloon_Adjust"), m_bAdjust);
*/

//注册表
class CPhdRegedit
{
	//注册表  主要分为根键(hKeyParent)、子键(lpszKeyName)和键值项三部分   
	//键值项：它又由三部分组成，分别为：名称(lpszValueName)、类型、数据(lpszValue)
	//hKeyParent：	HKEY_CLASSES_ROOT		HKEY_CURRENT_USER		HKEY_LOCAL_MACHINE		HKEY_USERS
public:
	CPhdRegedit(HKEY hKeyParent,LPCTSTR lpszKeyName);
	~CPhdRegedit();

	//************************************
	// Summary: 	设置键值项的值
	// Parameters: 	
	//    lpszValueName		- 	输入键值项的名称
	//    lpszValue		- 	输入键值项的值
	// Returns:   	
	//************************************
	bool SetValue(LPCTSTR lpszValueName,LPCTSTR lpszValue) const;
	bool SetValue(LPCTSTR lpszValueName, DWORD dwValue) const;
	
	//************************************
	// Summary: 	得到键值项的值
	// Parameters: 	
	//    lpszValueName		- 	输入键值项的名称
	//    szValue		- 	输出键值项的值
	// Returns:   	
	//************************************
	bool GetValue(LPCTSTR lpszValueName, CString& strValue) const;
	bool GetValue(LPCTSTR lpszValueName, DWORD& dwValue) const;

	//************************************
	// Summary: 	删除子键项
	// Parameters: 	
	//    lpszKeyName		- 	输入子键项子级的名称（直接输入名字，而不是路径）
	// Returns:   	
	//************************************
	bool DeleteChildKeyName(LPCTSTR lpszChildKeyName) const;

	//************************************
	// Summary: 	删除键值项
	// Parameters: 	
	//    lpszValueName		- 	输入键值项的名称
	// Returns:   	
	//************************************
	bool DeleteValue(LPCTSTR lpszValueName) const;

	//************************************
	// Summary: 	得到子键项下所有的子项名
	// Parameters: 	
	//    vecStrName		- 	输出子项名集合
	// Returns:   	
	//************************************
	bool GetAllChildKeyName(std::vector<CString>& vecStrName) const;

protected:
	HKEY m_hKeyParent; 
	LPCTSTR m_lpszKeyName;
	
};

