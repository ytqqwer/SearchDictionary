#include "stdafx.h"
#include "util.h"

std::string wstringToString(const std::wstring& wstr)
{
	LPCWSTR pwszSrc = wstr.c_str();
	int nLen = WideCharToMultiByte(CP_UTF8, 0, pwszSrc, -1, NULL, 0, NULL, NULL);
	if (nLen == 0)
		return std::string("");

	char* pszDst = new char[nLen];
	if (!pszDst)
		return std::string("");

	WideCharToMultiByte(CP_UTF8, 0, pwszSrc, -1, pszDst, nLen, NULL, NULL);
	std::string str(pszDst);
	delete[] pszDst;
	pszDst = NULL;

	return str;
}

std::wstring stringToWstring(const std::string& str)
{
	LPCSTR pszSrc = str.c_str();
	int nLen = MultiByteToWideChar(CP_UTF8, 0, pszSrc, -1, NULL, 0);
	if (nLen == 0)
		return std::wstring(L"");

	wchar_t* pwszDst = new wchar_t[nLen];
	if (!pwszDst)
		return std::wstring(L"");

	MultiByteToWideChar(CP_UTF8, 0, pszSrc, -1, pwszDst, nLen);
	std::wstring wstr(pwszDst);
	delete[] pwszDst;
	pwszDst = NULL;

	return wstr;
}

void TcharToChar(const TCHAR * tchar, char * _char)
{
	int iLength;
	//获取字节长度   
	iLength = WideCharToMultiByte(CP_UTF8, 0, tchar, -1, NULL, 0, NULL, NULL);
	//将tchar值赋给_char    
	WideCharToMultiByte(CP_UTF8, 0, tchar, -1, _char, iLength, NULL, NULL);
}

void CharToTchar(const char * _char, TCHAR * tchar)
{
	int iLength;
	iLength = MultiByteToWideChar(CP_UTF8, 0, _char, strlen(_char) + 1, NULL, 0);
	MultiByteToWideChar(CP_UTF8, 0, _char, strlen(_char) + 1, tchar, iLength);
}


void SplitString(const std::string& str, std::vector<std::string>& vector, const std::string& character)
{
	std::string::size_type pos1, pos2;
	pos2 = str.find(character);
	pos1 = 0;
	while (std::string::npos != pos2)
	{
		vector.push_back(str.substr(pos1, pos2 - pos1));

		pos1 = pos2 + character.size();
		pos2 = str.find(character, pos1);
	}
	if (pos1 != str.length())
		vector.push_back(str.substr(pos1));
}

void SplitWString(const std::wstring& str, std::vector<std::wstring>& vector, const std::wstring& character)
{
	std::wstring::size_type pos1, pos2;
	pos2 = str.find(character);
	pos1 = 0;
	while (std::wstring::npos != pos2)
	{
		vector.push_back(str.substr(pos1, pos2 - pos1));

		pos1 = pos2 + character.size();
		pos2 = str.find(character, pos1);
	}
	if (pos1 != str.length())
		vector.push_back(str.substr(pos1));
}
