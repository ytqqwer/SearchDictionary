#include "stdafx.h"
#include "ExcelReader.h"

ExcelReader::ExcelReader()
{
}

ExcelReader::~ExcelReader()
{
}

void ExcelReader::addXlsxFile(const std::string & fileName, const std::string& filePath)
{
	fileNames.push_back(fileName);

	XlsxFile xf(fileName, filePath);

	loadedXlsxFile.push_back(xf);
}

void ExcelReader::clear()
{
	fileNames.clear();
	loadedXlsxFile.clear();

	//selectedFileAndRows.clear();
}

std::vector<std::string> ExcelReader::getFileNames()
{
	return fileNames;
}

//std::string 
//ExcelReader::getValueInColumnByFileAndRow(const std::string & fileName, unsigned int row, const std::string & columnName)
//{
//
//	return std::string();
//}

void ExcelReader::findWord(const std::string& fileName, const std::string & word)
{
	//selectedFileAndRows.clear();
	//
	////检查全部文件
	//for (auto& xf : loadedXlsxFile) {

	//	bool isFound = false;
	//	std::vector<unsigned int> matchedRow;
	//	
	//	for (unsigned int curRow = 1; curRow < xf.size() - 1 ; curRow ++) {
	//		if (xf.getWordByRow(curRow) == word) {
	//			matchedRow.push_back(curRow);
	//			isFound = true;
	//		}
	//		else {
	//			continue;
	//		}
	//	}
	//	if (isFound) {
	//		selectedFileAndRows.push_back(make_pair(xf.getFileName(),matchedRow));
	//	}
	//}
	//
	//return selectedFileAndRows;

	//如果文件名为空，搜索全部文件
	if (fileName == "") {
		for (auto& xf : loadedXlsxFile) {
			xf.fineWord(word);
		}
	}
	else {
		for (auto& xf : loadedXlsxFile) {
			if (fileName == xf.getFileName()) {
				xf.fineWord(word);
				return;
			}
		}
	}	
}

void ExcelReader::setWordColumnName(const std::string & fileName, const std::string & columnName)
{
	for (auto& xf: loadedXlsxFile) {
		if (xf.getFileName() == fileName) {
			xf.setWordColumnName(columnName);
			return;
		}
	}
}

void ExcelReader::setDisplayColumnNames(const std::string & fileName, const std::vector<std::string>& displayColumnNames)
{
	for (auto& xf : loadedXlsxFile) {
		if (xf.getFileName() == fileName) {
			xf.setDisplayColumnNames(displayColumnNames);
			return;
		}
	}
}

std::string ExcelReader::getWordColumnName(const std::string & fileName)
{
	for (auto& xf : loadedXlsxFile) {
		if (xf.getFileName() == fileName) {
			return xf.getWordColumnName();
		}
	}
	return std::string();
}

std::vector<std::string> ExcelReader::getColumnNames(const std::string & fileName)
{
	for (auto& xf: loadedXlsxFile) {
		if (xf.getFileName() == fileName) {
			return xf.getColumnNames();
		}
	}
	return std::vector<std::string>();
}

std::vector<std::string> ExcelReader::getDisplayColumnNames(const std::string & fileName)
{
	for (auto& xf : loadedXlsxFile) {
		if (xf.getFileName() == fileName) {
			return xf.getDisplayColumnNames();
		}
	}
	return std::vector<std::string>();
}
