#include "StdAfx.h"
#include "CPhdConfig.h"

#define MAX_ALLKEYS 6000  //ȫ���ļ���
#define MAX_KEY 260  //һ����������
#define MAX_ALLSECTIONS 2048  //ȫ���Ķ���
#define MAX_SECTION 260  //һ����������


CPhdConfig::CPhdConfig(const CString& strFilePath)
	:m_strFilePath(strFilePath)
{
	CString strFileName = m_strFilePath.Right(m_strFilePath.GetLength() - m_strFilePath.ReverseFind(_T('\\')) - 1);
	m_strAppdataFile = GetAppdataDir() + _T("\\") + strFileName;
}

CPhdConfig::CPhdConfig()
{
	
}

CPhdConfig::~CPhdConfig()
{
}

void CPhdConfig::SetFilePath(LPCTSTR szFilePath)
{
	m_strFilePath = szFilePath;
	CString strFileName = m_strFilePath.Right(m_strFilePath.GetLength() - m_strFilePath.ReverseFind(_T('\\')) - 1);
	m_strAppdataFile = GetAppdataDir() + _T("\\") + strFileName;
}

bool CPhdConfig::SetValue(LPCTSTR szSection, LPCTSTR szName, LPCTSTR szValue) const
{
	return SetIniValue(szSection, szName, szValue);
}

bool CPhdConfig::SetValue(LPCTSTR szSection, LPCTSTR szName, int value) const
{
	TCHAR sz[10];
	_itot_s(value, sz, 10);
	return SetIniValue(szSection, szName, sz);
}

bool CPhdConfig::SetValue(LPCTSTR szSection, LPCTSTR szName, double value) const
{
	CString strValue;
	strValue.Format(_T("%G"), value);
	return SetIniValue(szSection, szName, strValue);
}

CString CPhdConfig::GetValue(LPCTSTR szSection, LPCTSTR szName, LPCTSTR szDefault /*= nullptr*/) const
{
	return GetIniValue(szSection, szName, szDefault);
}

int CPhdConfig::GetValue(LPCTSTR szSection, LPCTSTR szName, int nDefault) const
{
	TCHAR sz[10];
	_itot_s(nDefault, sz, 10);
	const auto result = GetIniValue(szSection, szName, sz);
	return _ttoi(result);
}

double CPhdConfig::GetValue(LPCTSTR szSection, LPCTSTR szName, double fDefault) const
{
	CString strDefault;
	strDefault.Format(_T("%G"), fDefault);

	const auto result = GetIniValue(szSection, szName, strDefault);
	return _ttof(result);
}

std::vector<CString> CPhdConfig::GetAllSections() const
{
	/*
	������������
	GetPrivateProfileSectionNames - �� ini �ļ��л�� Section ������
	��� ini �������� Section: [sec1] �� [sec2]���򷵻ص��� 'sec1',0,'sec2',0,0 �����㲻֪��
	ini ������Щ section ��ʱ���������� api ����ȡ����
	*/
	std::vector<CString> vecSection;

	int i;
	int iPos = 0;
	int iMaxCount;
	TCHAR chSectionNames[MAX_ALLSECTIONS] = { 0 }; //�ܵ���������ַ���
	TCHAR chSection[MAX_SECTION] = { 0 }; //���һ��������
	GetPrivateProfileSectionNames(chSectionNames, MAX_ALLSECTIONS, m_strFilePath);

	//����ѭ�����ضϵ�����������0
	for (i = 0; i < MAX_ALLSECTIONS; i++)
	{
		if (chSectionNames[i] == 0)
			if (chSectionNames[i] == chSectionNames[i + 1])
				break;
	}

	iMaxCount = i + 1; //Ҫ��һ��0��Ԫ�ء����ҳ�ȫ���ַ����Ľ������֡�

	for (i = 0; i < iMaxCount; i++)
	{
		chSection[iPos++] = chSectionNames[i];
		if (chSectionNames[i] == 0)
		{
			if (chSection != _T(""))
			{
				vecSection.push_back(chSection);
			}
			memset(chSection, 0, MAX_SECTION);
			iPos = 0;
		}
	}

	return vecSection;
}

BOOL CPhdConfig::DelSection(LPCTSTR lpSection) const
{
	return WritePrivateProfileString(lpSection, NULL, NULL, m_strFilePath);
}

BOOL CPhdConfig::DelAllSections() const
{
	std::vector<CString> vecSection = GetAllSections();
	for (int i = 0; i < vecSection.size(); i++)
		DelSection(vecSection[i]);
	
	return true;
}

BOOL CPhdConfig::DelKey(LPCTSTR lpSection, LPCTSTR lpKey) const
{
	return WritePrivateProfileString(lpSection, lpKey, NULL, m_strFilePath);
}

int CPhdConfig::GetSectionData(LPCTSTR lpszSection, std::vector<CString>& vecKey, std::vector<CString>& vecValue) const
{
	/*
	������������
	GetPrivateProfileSection- �� ini �ļ��л��һ��Section��ȫ��������ֵ��
	���ini����һ���Σ������� "��1=ֵ1" "��2=ֵ2"���򷵻ص��� '��1=ֵ1',0,'��2=ֵ2',0,0 �����㲻֪��
	���һ�����е����м���ֵ�����������
	*/
	int i = 0;
	int iPos = 0;
	CString strKeyValue;
	int iMaxCount = 0;
	TCHAR chKeyNames[MAX_ALLKEYS] = { 0 }; //�ܵ���������ַ���
	TCHAR chKey[MAX_KEY] = { 0 }; //�������һ������

	GetPrivateProfileSection(lpszSection, chKeyNames, MAX_ALLKEYS, m_strFilePath);

	for (i = 0; i < MAX_ALLKEYS; i++)
	{
		if (chKeyNames[i] == 0)
			if (chKeyNames[i] == chKeyNames[i + 1])
				break;
	}

	iMaxCount = i + 1; //Ҫ��һ��0��Ԫ�ء����ҳ�ȫ���ַ����Ľ������֡�

	for (i = 0; i < iMaxCount; i++)
	{
		chKey[iPos++] = chKeyNames[i];
		if (chKeyNames[i] == 0)
		{
			strKeyValue = chKey;
			CString strKey = strKeyValue.Left(strKeyValue.Find(_T('=')));
			if (strKey != _T(""))
			{
				vecKey.push_back(strKeyValue.Left(strKeyValue.Find(_T('='))));
				vecValue.push_back(strKeyValue.Mid(strKeyValue.Find(_T('=')) + 1));
			}
			memset(chKey, 0, MAX_KEY);
			iPos = 0;
		}
	}

	return (int)vecKey.size();
}

CString CPhdConfig::GetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR sz_default) const
{
	CString value;
	//��ȥAppdataĿ¼����
	auto dwRs = ::GetPrivateProfileString(section, valueName, sz_default, value.GetBuffer(256), 256, m_strAppdataFile);
	value.ReleaseBuffer();
	if (0x2 != GetLastError())
	{
		// �ɹ�
		return value;
	}
	//���appdataĿ¼��û�У����ڱ���Ŀ¼����
	dwRs = ::GetPrivateProfileString(section, valueName, sz_default, value.GetBuffer(256), 256, m_strFilePath);
	value.ReleaseBuffer();
	if (0x2 != GetLastError())
	{
		//�ɹ�
		return value;
	}
	//�����û�У���ѡ��Ĭ��ֵ
	return value = sz_default;
}

bool CPhdConfig::SetIniValue(LPCTSTR section, LPCTSTR valueName, LPCTSTR value) const
{
	bool bFlag = ::WritePrivateProfileString(section, valueName, value, m_strAppdataFile);
	if (!bFlag)
	{
		auto err = GetLastError();
	}
	return bFlag;
}


CString CPhdConfig::GetAppdataDir() const
{
	TCHAR szPath[_MAX_PATH];
	::SHGetSpecialFolderPath(nullptr, szPath, CSIDL_APPDATA, TRUE);
	CString strDir = szPath;
	strDir += _T("\\ZwSoftIni");
	::CreateDirectory(strDir, nullptr);

	return strDir;
}





