#pragma once

#include <string>
#include <vector>

#define WTS_U8(wstr) wstringToString(wstr)
#define STW_U8(str) stringToWstring(str)

std::string wstringToString(const std::wstring& wstr);
std::wstring stringToWstring(const std::string& str);

void TcharToChar(const TCHAR * tchar, char * _char);//��TCHARתΪchar��*tchar��TCHAR����ָ�룬*_char��char����ָ��   

void CharToTchar(const char * _char, TCHAR * tchar);//��charתΪTCHAR

void SplitString(const std::string& str, std::vector<std::string>& vector, const std::string& character);
void SplitWString(const std::wstring& str, std::vector<std::wstring>& vector, const std::wstring& character);