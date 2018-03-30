#pragma once

#include "xlnt/xlnt.hpp"

class XlsxFile
{
public:
	XlsxFile(const std::string& fileName, const std::string& path);
	~XlsxFile();

	void setWordColumnName(const std::string& wordColumnName);
	void setDisplayColumnNames(const std::vector<std::string>& displayColumnNames);
	
	//std::string getWordByRow(unsigned int);

	std::string getFileName();
	std::string getWordColumnName();
	std::vector<std::string> getColumnNames();
	std::vector<std::string> getDisplayColumnNames();

	unsigned int size();

	void fineWord(const std::string& word);

private:
	std::string fileName;
	std::string filePath;

	xlnt::workbook workbook;

	std::vector<std::string> columnNames;//全部列名
	std::vector<std::string> displayColumnNames;//要显示的列的列名
	//std::vector<std::pair<std::string, xlnt::cell_vector>> selectedColumns;	//已选择的列

	std::string wordColumnName;//“词语”所在列的列名
	std::vector<std::string> selectedWord;	//提取出的所有词语

	unsigned int maxRow;
};
