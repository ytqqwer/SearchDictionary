#include "stdafx.h"
#include "XlsxFile.h"

#include "common.h"

extern std::vector<FileNameAndColumnNames> FNACN;
extern std::vector<ListViewRow> listViewRows;

XlsxFile::XlsxFile(const std::string& name, const std::string& path)
{
	fileName = name;
	filePath = path;	
	workbook.load(path + name);

	//�����������������������
	//����length����Ϊ���ܻ�������
	xlnt::worksheet& ws = workbook.active_sheet();
	auto& rows = ws.rows(false);
	unsigned int i = 0;
	for (auto& row : rows) {
		//if (row[0].to_string() != u8"") {
		//	i++;
		//}
		i++;
	}
	maxRow = i;

	//��¼��������
	auto row = rows[0];
	for (auto& name : row) {
		std::string str = name.to_string();
		columnNames.push_back(str);
	}

}

XlsxFile::~XlsxFile()
{
}

void XlsxFile::setWordColumnName(const std::string & name)
{
	selectedWord.clear();

	wordColumnName = name;

	xlnt::worksheet& ws = workbook.active_sheet();
	auto& columns = ws.columns(false);
	for (auto& column : columns) {
			std::string str = column[0].to_string();

			//ʹ��xLnt��ȡxlsx�ļ�������ֵ��Ϊutf-8����
			//��str��ʵ�ʴ洢����utf-8������ַ���
			if (str == wordColumnName) {
				for (auto& word : column) {
					selectedWord.push_back(word.to_string());
				}
				return;
			}
	}
	
}

void XlsxFile::setDisplayColumnNames(const std::vector<std::string>& names)
{
	displayColumnNames = names;

	//����ѡ��֮ǰ������ա�
	//selectedColumns.clear();
	
	//xlnt::worksheet& ws = workbook.active_sheet();
	//auto columns = ws.columns(false);
	//for (auto& columnName : displayColumnNames) {
	//	for (auto& column : columns) {
	//		std::string str = column[0].to_string();

	//		//ʹ��xLnt��ȡxlsx�ļ�������ֵ��Ϊutf-8����
	//		//��str��ʵ�ʴ洢����utf-8������ַ���
	//		if (columnName == str) {
	//			selectedColumns.push_back(make_pair(columnName, column));
	//			break;
	//		}
	//	}
	//}
}

//std::string XlsxFile::getWordByRow(unsigned int index)
//{
//	auto& str = selectedWord[index];
//	return str;
//}

std::string XlsxFile::getFileName()
{
	return fileName;
}

std::string XlsxFile::getWordColumnName()
{
	return wordColumnName;
}

std::vector<std::string> XlsxFile::getColumnNames()
{
	//std::vector<std::string> names;

	//xlnt::worksheet& ws = workbook.active_sheet();
	//auto rows = ws.rows(false);
	//auto row = rows[0];
	//for (auto& name : row) {
	//	std::string str = name.to_string();
	//	names.push_back(str);
	//}

	return columnNames;
}

std::vector<std::string> XlsxFile::getDisplayColumnNames()
{
	return displayColumnNames;
}

unsigned int XlsxFile::size()
{
	return maxRow;
}

void XlsxFile::fineWord(const std::string & word)
{	
	xlnt::worksheet& ws = workbook.active_sheet();
	auto& rows = ws.rows(false);
	auto& firstRow = rows[0];

	//���׸���Ч�п�ʼ����
	for (unsigned int curRow = 1; curRow < maxRow - 1; curRow++) 
	{
		if (selectedWord[curRow] == word) {			
			auto& row = rows[curRow];
			ListViewRow lvr;
			lvr.fileName = fileName;

			unsigned int size = columnNames.size();
			for (unsigned int index = 0; index < size;index++) 
			{
				ListViewSubItem lvsi;
				lvsi.columnName = firstRow[index].to_string();
				lvsi.content = row[index].to_string();
				lvr.subItems.push_back(lvsi);
			}
			listViewRows.push_back(lvr);

		}
		else {
			continue;
		}
	}

}
