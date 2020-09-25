#pragma once
/***********************************************
   >   Class Name: PhdConfig
   >     Describe: 对ini配置文件可读写的封装类
   >       Author: peihaodong
   > Created Time: 2020年9月5日 
   >         Blog: https://blog.csdn.net/phd17621680432
 **********************************************/

class PhdConfig
{
	/*配置文件主要分为三部分：字段名(section)	键名(key)	键值(value)*/
public:
	PhdConfig();
	PhdConfig(const CString& strFilePath);
	~PhdConfig();

	//************************************
	// Summary: 	设置文件路径
	// Parameters: 	
	//    szFilePath		- 	
	// Returns:   	
	//************************************
	void SetFilePath(LPCTSTR szFilePath);

	//************************************
	// Summary: 	写入键值
	// Parameters: 	
	//    szSection		- 	输入字段名
	//    szName		- 	输入键名
	//    szValue		- 	输入键值
	// Returns:   	
	//************************************
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, LPCTSTR szValue) const;
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, int value) const;
	bool SetValue(LPCTSTR szSection, LPCTSTR szKey, double value) const;

	//************************************
	// Summary: 	读取键值
	// Parameters: 	
	//    szSection		- 	输入字段名
	//    szName		- 	输入键名
	//    szDefault		- 	输入默认值（如果配置文件中没有读到，则返回默认值）
	// Returns:   	
	//************************************
	CString GetValue(LPCTSTR szSection, LPCTSTR szKey, LPCTSTR szDefault = nullptr) const;
	int GetValue(LPCTSTR szSection, LPCTSTR szKey, int nDefault) const;
	double GetValue(LPCTSTR szSection, LPCTSTR szKey, double fDefault) const;

	//************************************
	// Summary: 	得到该配置文件的所有字段名
	// Parameters: 	
	// Returns:   	
	//************************************
	std::vector<CString> GetAllSections() const;

	//************************************
	// Summary: 	删除某一字段
	// Parameters: 	
	//    lpSection		- 	输入字段名
	// Returns:   	
	//************************************
	BOOL DelSection(LPCTSTR lpSection) const;

	//************************************
	// Summary: 	删除所有字段
	// Parameters: 	
	// Returns:   	
	//************************************
	BOOL DelAllSections() const;

	//************************************
	// Summary: 	删除键名
	// Parameters: 	
	//    lpSection		- 	输入字段名
	//    lpKey		- 	输入键名
	// Returns:   	
	//************************************
	BOOL DelKey(LPCTSTR lpSection, LPCTSTR lpKey) const;

	//************************************
	// Summary: 	得到字段下所有的键名和键值
	// Parameters: 	
	//    lpszSection		- 	输入字段名
	//    vecKey		- 	输出键名集合
	//    vecValue		- 	输出键值集合
	// Returns:   	返回字段下键的总数
	//************************************
	int GetSectionData(LPCTSTR lpszSection, std::vector<CString>& vecKey, std::vector<CString>& vecValue) const;

protected:
	CString m_strFilePath;
	CString m_strAppdataFile;	// 放置在appdata目录下

	CString GetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR sz_default) const;
	bool SetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR value) const;

private:
	CString GetAppdataDir() const;
};