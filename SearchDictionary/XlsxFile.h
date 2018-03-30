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

	std::vector<std::string> columnNames;//ȫ������
	std::vector<std::string> displayColumnNames;//Ҫ��ʾ���е�����
	//std::vector<std::pair<std::string, xlnt::cell_vector>> selectedColumns;	//��ѡ�����

	std::string wordColumnName;//����������е�����
	std::vector<std::string> selectedWord;	//��ȡ�������д���

	unsigned int maxRow;
};
