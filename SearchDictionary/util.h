#pragma once

#include <string>
#include <vector>

#define WTS_U8(wstr) wstringToString(wstr)
#define STW_U8(str) stringToWstring(str)

std::string wstringToString(const std::wstring& wstr);
std::wstring stringToWstring(const std::string& str);

void TcharToChar(const TCHAR * tchar, char * _char);//将TCHAR转为char，*tchar是TCHAR类型指针，*_char是char类型指针   

void CharToTchar(const char * _char, TCHAR * tchar);//把char转为TCHAR

void SplitString(const std::string& str, std::vector<std::string>& vector, const std::string& character);
void SplitWString(const std::wstring& str, std::vector<std::wstring>& vector, const std::wstring& character);