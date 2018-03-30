#pragma once

#include <string>

class IniFile
{
public:
	IniFile(LPCTSTR szFileName);
	~IniFile();

public:
	// Attributes     
	void SetFileName(LPCTSTR szFileName);

public:
	// Operations  
	//DWORD GetProfileSectionNames(CStringArray& strArray); // ����section����  

	//LPCTSTR ---> LP Const Tchar Str
	BOOL SetProfileString(LPCTSTR lpszSectionName, LPCTSTR lpszKeyName, LPCTSTR lpszKeyValue);
	DWORD GetProfileString(LPCTSTR lpszSectionName, LPCTSTR lpszKeyName, LPTSTR szKeyValue);
	
	//BOOL SetProfileInt(LPCTSTR lpszSectionName, LPCTSTR lpszKeyName, int nKeyValue);
	//int GetProfileInt(LPCTSTR lpszSectionName, LPCTSTR lpszKeyName);

	BOOL DeleteSection(LPCTSTR lpszSectionName);
	BOOL DeleteKey(LPCTSTR lpszSectionName, LPCTSTR lpszKeyName);

private:
	std::wstring  m_szFileName; // .//Config.ini, ������ļ������ڣ���exe��һ����ͼWriteʱ���������ļ�  

	UINT m_unMaxSection; // ���֧�ֵ�section��(256)  
	UINT m_unSectionNameMaxSize; // section���Ƴ��ȣ�������Ϊ32(Null-terminated)  
};
