#pragma once

#include "XlsxFile.h"

class ExcelReader
{
public:
	ExcelReader();
	~ExcelReader();

	void addXlsxFile(const std::string& fileName, const std::string& path);

	void clear();

	void findWord(const std::string& fileName,const std::string& word);

	//std::string getValueInColumnByFileAndRow(const std::string& fileName, unsigned int row, const std::string& columnName);

public:
	void setWordColumnName(const std::string& fileName , const std::string& wordColumnName);
	void setDisplayColumnNames(const std::string& fileName, const std::vector<std::string>& displayColumnNames);
	
	std::string getWordColumnName(const std::string& fileName);
	std::vector<std::string> getFileNames();
	std::vector<std::string> getColumnNames(const std::string& fileName);
	std::vector<std::string> getDisplayColumnNames(const std::string& fileName);

private:
	std::vector<std::string> fileNames;
	std::vector<XlsxFile> loadedXlsxFile;	 //已加载的工作簿，对应一个文件名

	//std::vector<std::pair<const std::string , std::vector<unsigned int>>> selectedFileAndRows;	//已选择的文件名与行
};
