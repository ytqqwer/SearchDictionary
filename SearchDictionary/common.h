#pragma once

#include <string>
#include <vector>

struct FileNameAndColumnNames {
	std::string fileName;
	std::vector<std::string> columnNames;
	HWND groupBox1;
	HWND groupBox2;
	std::vector<HWND> checkBoxHandles;
	std::vector<HWND> radioButtonHandles;
};

struct ListViewSubItem {
	std::string columnName;
	std::string content;
};

struct ListViewRow {
	std::vector<ListViewSubItem> subItems;
	std::string fileName;
};
