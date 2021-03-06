#include "stdafx.h"
#include "SearchDictionary.h"

#include <stdio.h>
#include <commdlg.h>
#include <CommCtrl.h>

#include <ShlObj.h>			//选择文件夹用

#include <codecvt>
#include <io.h>				//遍历文件使用

#include "ExcelReader.h"
#include "IniFile.h"

#include "util.h"
#include "common.h"

#define MAX_LOADSTRING 100

std::vector<FileNameAndColumnNames> FNACN;

std::vector<ListViewRow> listViewRows;

// Global Variables: 
HINSTANCE hInst;                                // current instance

//////////////////////////////////////////////////////////////////////////////////
//主窗口
WCHAR szTitle[MAX_LOADSTRING];                  // 标题栏文本

WCHAR szMainWindowClass[MAX_LOADSTRING];        // 主窗口类名

HWND hMainWindow;							// 主窗口句柄

HWND hSearchEdit;							// 搜索框句柄
HWND hSearchButton;							// 搜索按钮句柄

HWND hSettingButton;						// 设置按钮句柄

HWND hFileNameComboBox;						// 文件名下拉列表句柄
HWND hDictionaryListView;					// 词典列表视图句柄

//////////////////////////////////////////////////////////////////////////////////
//词语详情窗口
WCHAR szDetailModelWindowClass[MAX_LOADSTRING];       // 词语详情模态窗口类名

HWND hDetailModelWindow;							// 词语详情模态窗口句柄

//////////////////////////////////////////////////////////////////////////////////
//设置窗口
WCHAR szSettingModelWindowClass[MAX_LOADSTRING];	//参数设置模态窗口类名

HWND hSettingModelWindow;							// 设置窗口句柄

HWND hFileNameListView;					    // 文件名列表视图句柄

HWND hSettingModelWindowSaveButton;
HWND hSettingModelWindowCancelButton;

//////////////////////////////////////////////////////////////////////////////////
WNDPROC oldSearchEditProc;					// 暂存原搜索编辑框处理过程

ExcelReader* reader;
IniFile* iniFile;

// Foward declarations of functions included in this code module:
ATOM                RegisterMyClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);

INT_PTR CALLBACK    AboutProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK    MainWindowProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK    DetailModelWindowProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK    SettingModelWindowProc(HWND, UINT, WPARAM, LPARAM);

VOID ModalWindowMessageLoop(HWND);
VOID DisplayDetailModelWindow(HWND,unsigned int);
VOID DisplaySettingModelWindow(HWND);
VOID ActiveControls(bool);
VOID OpenFiles();
VOID SetSelectedColumnNamesAndWordColumnName();
VOID refreshListView();
VOID Search();

//搜索编辑框处理过程
LRESULT CALLBACK	SearchEditProc(HWND, UINT, WPARAM, LPARAM);

///////////////////////////////////////////////////////////////////////////////////////

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// Initialize global strings
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_MAIN_WINDOW, szMainWindowClass, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_DETAIL_MODEL_WINDOW, szDetailModelWindowClass, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_SETTING_MODEL_WINDOW, szSettingModelWindowClass, MAX_LOADSTRING);

	RegisterMyClass(hInstance);

	// Perform application initialization:
	if (!InitInstance(hInstance, nCmdShow))
	{
		return FALSE;
	}
	
	reader = new ExcelReader();
	iniFile = new IniFile(L"config.ini");

	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_MAIN_WINDOW));
	
	// Main message loop:
	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	delete reader;
	delete iniFile;

	return (int)msg.wParam;
}

//
//  函数: RegisterMyClass()
//
//  目的: 注册窗口类。
//
ATOM RegisterMyClass(HINSTANCE hInstance)
{
	WNDCLASSEX Main_Window_Class_Ex;
	WNDCLASSEX Detail_Model_Window_Class_Ex;
	WNDCLASSEX Setting_Model_Window_Class_Ex;

	Main_Window_Class_Ex.cbSize = sizeof(WNDCLASSEX);
	Main_Window_Class_Ex.style = CS_HREDRAW | CS_VREDRAW;
	Main_Window_Class_Ex.cbClsExtra = 0;
	Main_Window_Class_Ex.cbWndExtra = 0;
	Main_Window_Class_Ex.hInstance = hInstance;
	Main_Window_Class_Ex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SEARCHDICTIONARY));
	Main_Window_Class_Ex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	Main_Window_Class_Ex.hbrBackground = (HBRUSH)(COLOR_WINDOW);
	Main_Window_Class_Ex.lpszMenuName = MAKEINTRESOURCEW(IDC_MAIN_WINDOW);
	Main_Window_Class_Ex.lpfnWndProc = MainWindowProc;
	Main_Window_Class_Ex.lpszClassName = szMainWindowClass;
	Main_Window_Class_Ex.hIconSm = LoadIcon(Main_Window_Class_Ex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	Detail_Model_Window_Class_Ex.cbSize = sizeof(WNDCLASSEX);
	Detail_Model_Window_Class_Ex.style = CS_HREDRAW | CS_VREDRAW;
	Detail_Model_Window_Class_Ex.cbClsExtra = 0;
	Detail_Model_Window_Class_Ex.cbWndExtra = 0;
	Detail_Model_Window_Class_Ex.hInstance = hInstance;
	Detail_Model_Window_Class_Ex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SEARCHDICTIONARY));
	Detail_Model_Window_Class_Ex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	Detail_Model_Window_Class_Ex.hbrBackground = (HBRUSH)(COLOR_WINDOW);
	Detail_Model_Window_Class_Ex.lpszMenuName = NULL;
	Detail_Model_Window_Class_Ex.lpfnWndProc = DetailModelWindowProc;
	Detail_Model_Window_Class_Ex.lpszClassName = szDetailModelWindowClass;
	Detail_Model_Window_Class_Ex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));

	Setting_Model_Window_Class_Ex.cbSize = sizeof(WNDCLASSEX);
	Setting_Model_Window_Class_Ex.style = CS_HREDRAW | CS_VREDRAW;
	Setting_Model_Window_Class_Ex.cbClsExtra = 0;
	Setting_Model_Window_Class_Ex.cbWndExtra = 0;
	Setting_Model_Window_Class_Ex.hInstance = hInstance;
	Setting_Model_Window_Class_Ex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SEARCHDICTIONARY));
	Setting_Model_Window_Class_Ex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	Setting_Model_Window_Class_Ex.hbrBackground = (HBRUSH)(COLOR_WINDOW);
	Setting_Model_Window_Class_Ex.lpszMenuName = NULL;
	Setting_Model_Window_Class_Ex.lpfnWndProc = SettingModelWindowProc;
	Setting_Model_Window_Class_Ex.lpszClassName = szSettingModelWindowClass;
	Setting_Model_Window_Class_Ex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
	
	return (
		RegisterClassEx(&Main_Window_Class_Ex) && 
		RegisterClassEx(&Detail_Model_Window_Class_Ex) && 
		RegisterClassEx(&Setting_Model_Window_Class_Ex)
		);
}

//
//   函数: InitInstance(HINSTANCE, int)
//
//   目的: 保存实例句柄并创建主窗口
//
//   注释: 
//
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance; // 将实例句柄存储在全局变量中

	HWND hWnd = CreateWindowW(szMainWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, 1000, 450, nullptr, nullptr, hInstance, nullptr);

	hMainWindow = hWnd;

	if (!hWnd)
	{
		return FALSE;
	}

	//////////////////////////////////////////////////////////////////////
	//初始化搜索编辑框
	hSearchEdit = CreateWindow(_T("EDIT"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_LEFT,
		70, 20, 200, 30, hWnd, (HMENU)ID_SEARCH_EDIT, hInst, NULL);
	//设置处理过程
	oldSearchEditProc = (WNDPROC)SetWindowLongPtr(hSearchEdit, GWLP_WNDPROC, (LONG_PTR)SearchEditProc);
	
	//////////////////////////////////////////////////////////////////////
	//初始化搜索按钮
	hSearchButton = CreateWindow(_T("BUTTON"), _T("搜索"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		310, 20, 100, 30, hWnd, (HMENU)ID_SEARCH_BUTTON, hInst, NULL);

	//////////////////////////////////////////////////////////////////////
	//初始化设置按钮
	hSettingButton = CreateWindow(_T("BUTTON"), _T("设置"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		470, 20, 100, 30, hWnd, (HMENU)ID_SETTING_BUTTON, hInst, NULL);
	
	//////////////////////////////////////////////////////////////////////
	//初始化文本	
	HWND hWordText = CreateWindow(_T("static"), _T("词语"), WS_CHILD | WS_VISIBLE | SS_LEFT, 30, 25, 30, 30, hWnd,
		NULL, hInst, NULL);

	//////////////////////////////////////////////////////////////////////
	//初始化文件名下拉列表
	hFileNameComboBox = CreateWindow(WC_COMBOBOX, _T(""), CBS_DROPDOWNLIST | CBS_HASSTRINGS | WS_CHILD | WS_OVERLAPPED | WS_VISIBLE,
		600, 20, 160, 800, hWnd, (HMENU)ID_FILENAME_COMBOBOX, hInst, NULL);
	
	//////////////////////////////////////////////////////////////////////
	//初始化词典的列表视图
	//无需自动加载词语信息并显示
	hDictionaryListView = CreateWindowW(WC_LISTVIEW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_NOSORTHEADER ,
		25, 65, 940, 300, hWnd, (HMENU)ID_DICTIONARY_LISTVIEW, hInst, NULL);

	ListView_SetExtendedListViewStyle(hDictionaryListView, LVS_EX_FULLROWSELECT);     //设置整行选择风格
	
	//////////////////////////////////////////////////////////////////////
	//设置字体
	HFONT hFont = CreateFont(20, 0, 0, 0, 0, FALSE, FALSE, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"微软雅黑");//创建字体
	SendMessage(hSearchButton, WM_SETFONT, (WPARAM)hFont, TRUE);//发送设置字体消息
	SendMessage(hSearchEdit, WM_SETFONT, (WPARAM)hFont, TRUE); 
	SendMessage(hWordText, WM_SETFONT, (WPARAM)hFont, TRUE);
	SendMessage(hSettingButton, WM_SETFONT, (WPARAM)hFont, TRUE);
	SendMessage(hFileNameComboBox, WM_SETFONT, (WPARAM)hFont, TRUE);

	//////////////////////////////////////////////////////////////////////
	// 禁用某些功能，直到被激活，TRUE使用，FALSE禁止
	ActiveControls(FALSE);
		
	///////////////////////////////////////////////////////////////////////
	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	return TRUE;
}

//
//  函数: MainWindowProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的:    处理主窗口的消息。
//
//  WM_COMMAND  - 处理应用程序菜单
//  WM_PAINT    - 绘制主窗口
//  WM_DESTROY  - 发送退出消息并返回
//
LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{

	case WM_NOTIFY:
	{
		switch (wParam)
		{
		case ID_DICTIONARY_LISTVIEW:
		{
			LPNMITEMACTIVATE activeItem = (LPNMITEMACTIVATE)lParam; //得到NMITEMACTIVATE结构指针  
			if (activeItem->iItem >= 0)				//当点击空白区域时，iItem值为-1，防止越界
			{
				switch (activeItem->hdr.code) {		//判断通知码  
				case NM_CLICK: //单击
				{
					//do nothing

				}	break;
				case NM_DBLCLK: //双击
				{
					//char index[50];
					//_itoa_s(activeItem->iItem, index, 10);	//activeItem->iItem项目序号，将int整型数转化为一个字符串
					//MessageBoxA(hWnd, index, "双击", 0);

					DisplayDetailModelWindow(hMainWindow,activeItem->iItem);

				}	break;
				case NM_RCLICK: //右击
				{
					char index[50];
					_itoa_s(activeItem->iItem, index, 10);
					MessageBoxA(hWnd, index, "右击", 0);
				}   break;
				}
			}

		}break;
		}
	}break;

	case WM_COMMAND:
	{
		// 分析wParam高位: 
		switch (HIWORD(wParam))
		{
		case CBN_SELCHANGE://列表框选择发生变化
		{
			//do nothing
		}
		break;
		default:
			//继续处理
			break;
		}

		// 分析wParam低位: 
		switch (LOWORD(wParam))
		{
		case ID_SEARCH_BUTTON:
		{
			Search();
		}
		break;		
		case ID_SETTING_BUTTON:
		{
			DisplaySettingModelWindow(hMainWindow);
		}
		break;
		case IDM_OPEN:
		{
			OpenFiles();
		}
		break;
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, AboutProc);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}

	}
	break;
	case WM_PAINT:
	{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(hWnd, &ps);
		// TODO: 在此处添加使用 hdc 的任何绘图代码...
		EndPaint(hWnd, &ps);
	}
	break;
	case WM_DESTROY:
	{
		PostQuitMessage(0);
	}
	break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// 设置模态窗口的消息处理程序。
LRESULT CALLBACK SettingModelWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_NOTIFY:
		{
			switch (wParam)
			{
				case ID_FILENAME_LISTVIEW:
				{
					LPNMITEMACTIVATE activeItem = (LPNMITEMACTIVATE)lParam; //得到NMITEMACTIVATE结构指针  
					if (activeItem->iItem >= 0)				//当点击空白区域时，iItem值为-1，防止越界
					{
						switch (activeItem->hdr.code) {		//判断通知码  
							case NM_CLICK: //单击
							{
								for (auto& fncn : FNACN) {
									ShowWindow(fncn.groupBox1, SW_HIDE);
									ShowWindow(fncn.groupBox2, SW_HIDE);
								}
								ShowWindow(FNACN[activeItem->iItem].groupBox1, SW_SHOW);
								ShowWindow(FNACN[activeItem->iItem].groupBox2, SW_SHOW);

							}	break;
							case NM_DBLCLK: 
							{
								char a[50];
								_itoa_s(activeItem->iItem, a, 10);	//activeItem->iItem项目序号，将int整型数转化为一个字符串
								MessageBoxA(hWnd, a, "双击", 0);  
							}	break;
							case NM_RCLICK: 
							{
								char a[50];
								_itoa_s(activeItem->iItem, a, 10);
								MessageBoxA(hWnd, a, "右击", 0);  
							}   break;
						}
					}
	
				}break;
			}
		}break;

		case WM_COMMAND:
		{			
			switch (LOWORD(wParam))// 分析wParam低位
			{
				case ID_SETTING_MODEL_WINDOW_SAVE_BUTTON:
				{
					//保存选择的结果至文件
					for (auto& fncn :FNACN) {
						std::wstring wstr;
						TCHAR tchar[1000];

						for (auto handle : fncn.checkBoxHandles) {
							if (SendMessage(handle, BM_GETCHECK, 0, 0) == TRUE)//是否打勾了 
							{
								GetWindowText(handle,tchar,1000);
								std::wstring temp = tchar;
								wstr = wstr + L"," + temp;
							}
						}
						iniFile->SetProfileString(STW_U8(fncn.fileName).c_str(),L"SelectedColumnNames",wstr.c_str());

						for (auto handle : fncn.radioButtonHandles) {
							if (SendMessage(handle, BM_GETCHECK, 0, 0) == TRUE)
							{
								GetWindowText(handle, tchar, 1000);
								break;
							}
						}
						iniFile->SetProfileString(STW_U8(fncn.fileName).c_str(), L"WordColumnName", tchar);

					}

					SetSelectedColumnNamesAndWordColumnName();
					
					MessageBox(hWnd, L"保存", L"提示", MB_OK);
				}	break;
				case ID_SETTING_MODEL_WINDOW_CANCEL_BUTTON: 
				{
					DestroyWindow(hWnd);//销毁当前窗口
				}	break;
			}
		}break;
		case WM_DESTROY: {		//关闭窗口时总是处理
			PostQuitMessage(0);//发送消息，退出当前窗口的专用事件循环
			return 0;
		}break;
	}
	return DefWindowProc(hWnd, message, wParam, lParam);
}

// 搜索编辑框 的消息处理函数
// 处理回车事件专用
LRESULT CALLBACK SearchEditProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_RETURN:
		{
			Search();
		}
		break;  //or return 0; if you don't want to pass it further to def proc
				//If not your key, skip to default:
		}
	default:
		return CallWindowProc(oldSearchEditProc, hWnd, msg, wParam, lParam);
	}
	return 0;
}

// 详情模态窗口的消息处理程序。
LRESULT CALLBACK DetailModelWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message) {
	case WM_DESTROY:
		PostQuitMessage(0);
		return 0;
	}
	return DefWindowProc(hWnd, message, wParam, lParam);
}

// “关于”框的消息处理程序。
INT_PTR CALLBACK AboutProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

void ModalWindowMessageLoop(HWND last)
{
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	EnableWindow(last, TRUE);
	SetForegroundWindow(last);
}

VOID DisplayDetailModelWindow(HWND hParent,unsigned int rowIndex)
{
	EnableWindow(hParent, FALSE);

	//TODO 创建适应文本大小的窗口，以及调整输出格式

	auto& lvr = listViewRows[rowIndex];

	auto size = lvr.subItems.size();

	HWND hWnd = CreateWindowW(
		szDetailModelWindowClass,
		TEXT("Detail"),
		WS_OVERLAPPEDWINDOW | WS_VISIBLE,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		400,
		50 + size * 35,
		hParent,
		NULL,
		(HINSTANCE)GetWindowLong(hParent, GWL_HINSTANCE),
		NULL);

	hDetailModelWindow = hWnd;
	
	unsigned int i = 0;
	for (auto& si : lvr.subItems) 
	{
		HWND handle = CreateWindow(_T("static"), (LPWSTR)STW_U8(si.columnName).c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 
			10, 10 + 28 * i, 90, 30, hWnd,	NULL, hInst, NULL);
		SendMessage(handle, WM_SETFONT, (WPARAM)GetStockObject(17), 0);

		handle = CreateWindow(_T("static"), (LPWSTR)STW_U8(si.content).c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT,
			105, 10 + 28 * i, 260, 30, hWnd, NULL, hInst, NULL);
		SendMessage(handle, WM_SETFONT, (WPARAM)GetStockObject(17), 0);


		i++;
	}

	ShowWindow(hWnd, SW_SHOW);
	UpdateWindow(hWnd);

	ModalWindowMessageLoop(hParent);
}

VOID DisplaySettingModelWindow(HWND hParent)
{
	EnableWindow(hParent, FALSE);

	HWND hWnd = CreateWindowW(
		szSettingModelWindowClass,
		TEXT("Setting"),
		WS_OVERLAPPEDWINDOW | WS_VISIBLE,
		50,
		50,
		700,
		600,
		hParent,
		NULL,
		hInst,
		NULL);

	hSettingModelWindow = hWnd;

	hSettingModelWindowSaveButton = CreateWindow(_T("BUTTON"), _T("保存"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		420, 500, 100, 30, hWnd, (HMENU)ID_SETTING_MODEL_WINDOW_SAVE_BUTTON, hInst, NULL);

	hSettingModelWindowCancelButton = CreateWindow(_T("BUTTON"), _T("取消"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		550, 500, 100, 30, hWnd, (HMENU)ID_SETTING_MODEL_WINDOW_CANCEL_BUTTON, hInst, NULL);

	SendMessage(hSettingModelWindowSaveButton, WM_SETFONT, (WPARAM)GetStockObject(17), 0);
	SendMessage(hSettingModelWindowCancelButton, WM_SETFONT, (WPARAM)GetStockObject(17), 0);
	
	/////////////////////////////////////////////////////////////////////////
	//初始化文件名列表视图
	hFileNameListView = CreateWindowW(WC_LISTVIEW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_NOSORTHEADER,
		25, 25, 200, 500, hWnd, (HMENU)ID_FILENAME_LISTVIEW, hInst, NULL);

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvc.iSubItem = 0;
	lvc.pszText = TEXT("文件名");
	lvc.fmt = LVCFMT_LEFT;
	lvc.cx = 200;	
	ListView_InsertColumn(hFileNameListView, 0, &lvc);

	//清空，重新加载
	FNACN.clear();

	//加载文件名
	LVITEM vitem;
	vitem.mask = LVIF_TEXT;
	vitem.iItem = 0;
	vitem.iSubItem = 0;

	auto fileNames = reader->getFileNames();
	for (auto name : fileNames) {
		FileNameAndColumnNames fncn;
		fncn.fileName = name;
		fncn.columnNames = reader->getColumnNames(name);
		FNACN.push_back(fncn);

		std::wstring wstr;
		wstr = stringToWstring(name);
		vitem.pszText = (LPWSTR)wstr.c_str();
		ListView_InsertItem(hFileNameListView, &vitem);
		vitem.iItem ++;

	}

	/////////////////////////////////////////////////////////////////////////
	//对每一个文件名，创建一系列单选或多选框，使用GroupBox包装
	for (auto& fncn : FNACN) {
		
		int size = fncn.columnNames.size();
		int rowCount;
		if (size % 4) {
			rowCount = size / 4 + 1;
		}
		else {
			rowCount = size / 4;
		}
		
		//加载GroupBox
		HWND groupBox1 = CreateWindow(TEXT("Button"), TEXT("选择要显示的列"), WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
			250, 15, 410, rowCount * 30 + 15, 
			hWnd, NULL, hInst, NULL);	
		HWND groupBox2 = CreateWindow(TEXT("Button"), TEXT("选择词语所在的列"), WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
			250, 15 + rowCount * 30 + 30, 410 , rowCount * 30 + 15, 
			hWnd, NULL, hInst, NULL);

		SendMessage(groupBox1, WM_SETFONT, (WPARAM)GetStockObject(17), 0);
		SendMessage(groupBox2, WM_SETFONT, (WPARAM)GetStockObject(17), 0);
		fncn.groupBox1 = groupBox1;
		fncn.groupBox2 = groupBox2;
		
		//加载CheckBox
		for (int i = 0; i < size;i++) {

			auto name = fncn.columnNames[i];
			std::wstring wstr;
			wstr = stringToWstring(name);

			HWND checkBox = CreateWindow(TEXT("Button"), wstr.c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
				10 + i % 4 * 100, 10 + i / 4 * 30, 90, 30, 
				groupBox1, NULL, hInst, NULL);
			SendMessage(checkBox, WM_SETFONT, (WPARAM)GetStockObject(17), 0);

			fncn.checkBoxHandles.push_back(checkBox);
		}
		
		//加载RadioButton
		for (int i = 0; i < size; i++) {
			auto name = fncn.columnNames[i];
			std::wstring wstr;
			wstr = stringToWstring(name);

			HWND radioButton = CreateWindow(TEXT("Button"), wstr.c_str(), WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
				10 + i % 4 * 100, 10 + i / 4 * 30, 90, 30,
				groupBox2, NULL, hInst, NULL);
			SendMessage(radioButton, WM_SETFONT, (WPARAM)GetStockObject(17), 0);

			fncn.radioButtonHandles.push_back(radioButton);
		}

	}
	
	/////////////////////////////////////////////////////////////////////////
	//默认只显示一组GroupBox
	for (auto& fncn : FNACN) {
		ShowWindow(fncn.groupBox1,SW_HIDE);
		ShowWindow(fncn.groupBox2, SW_HIDE);
	}
	ShowWindow(FNACN[0].groupBox1, SW_SHOW);
	ShowWindow(FNACN[0].groupBox2, SW_SHOW);

	/////////////////////////////////////////////////////////////////////////
	// 按钮打勾
	for (auto& fncn : FNACN) 
	{
		TCHAR tchar[1000];
		std::vector<std::wstring> columnNames;

		//check box
		iniFile->GetProfileStringW(STW_U8(fncn.fileName).c_str(), L"SelectedColumnNames" , tchar);		
		SplitWString(tchar,columnNames,L",");
		
		for (auto& handle : fncn.checkBoxHandles) {
			GetWindowText(handle, tchar, 1000);
			std::wstring temp = tchar;

			for (auto& name : columnNames) 
			{
				if (temp == name) 
				{
					SendMessage(handle, BM_SETCHECK, BST_CHECKED, 0);
					break;
				}
			}
		}

		//radio button
		iniFile->GetProfileStringW(STW_U8(fncn.fileName).c_str(), L"WordColumnName", tchar);
		std::wstring wordColumnName = tchar;
		for (auto& handle : fncn.radioButtonHandles) {
			GetWindowText(handle, tchar, 1000);
			std::wstring temp = tchar;

			if (temp == wordColumnName)
			{
				SendMessage(handle, BM_SETCHECK, BST_CHECKED, 0);
				break;
			}
		}

	}

	ShowWindow(hWnd, SW_SHOW);
	UpdateWindow(hWnd);

	ModalWindowMessageLoop(hParent);
}


void ActiveControls(bool boolean)
{
	EnableWindow(hSearchButton, boolean);
	EnableWindow(hSearchEdit, boolean);
	EnableWindow(hFileNameComboBox, boolean);
	EnableWindow(hDictionaryListView, boolean);
	EnableWindow(hSettingButton, boolean);
}

VOID Search() 
{
	listViewRows.clear();

	///////////////////////////////////////////////////////////////////////////
	TCHAR buff[80] = _T("");
	GetWindowText(hSearchEdit, buff, 80);
	const std::string& word = WTS_U8(buff);

	//CB_GETCURSEL - 获取当前被选择项。
	//	- WPARAM / LPARAM 均未使用，必须为0
	//	- SendMessage函数的返回，获取选择项的索引，如果当前没有选择项返回CB_ERR
	int index = SendMessage(hFileNameComboBox, CB_GETCURSEL, 0, 0);

	//CB_GETLBTEXT - 获取选项字符的内容
	//	- WPARAM - 选项的索引
	//	- LPARAM - 获取字符内容的BUFF
	TCHAR fileName[200];
	SendMessage(hFileNameComboBox, CB_GETLBTEXT, index, (LPARAM)fileName);
	std::wstring wstr = fileName;

	reader->findWord(WTS_U8(wstr),word);

	///////////////////////////////////////////////////////////////////////////
	if (listViewRows.size()) {
		refreshListView();
	}
	else {
		MessageBox(hMainWindow, L"在当前文件下未找到该词语。", L"提示", MB_OK);	
	}

}

VOID SetSelectedColumnNamesAndWordColumnName() 
{
	auto fileNames = reader->getFileNames();
	for (auto& name : fileNames)
	{
		TCHAR tchar[1000];
		std::vector<std::wstring> columnNamesW;

		//SelectedColumnNames
		iniFile->GetProfileStringW(STW_U8(name).c_str(), L"SelectedColumnNames", tchar);
		SplitWString(tchar, columnNamesW, L",");
		std::vector<std::string> columnNames;
		for (auto & wstr : columnNamesW) {
			if (wstr.length() != 0) {
				columnNames.push_back(WTS_U8(wstr));
			}
		}
		reader->setDisplayColumnNames(name, columnNames);

		//WordColumnName
		iniFile->GetProfileStringW(STW_U8(name).c_str(), L"WordColumnName", tchar);
		std::wstring wordColumnName = tchar;
		reader->setWordColumnName(name, WTS_U8(wordColumnName));
	}

}

void refreshListView() 
{
	// 清除ListView中的所有项 
	//ListView_DeleteAllItems(hDictionaryListView);
	////删除所有行
	SendMessage(hDictionaryListView, LVM_DELETEALLITEMS, 0, 0);

	//得到ListView的Header窗体
	HWND hWndListViewHeader = (HWND)SendMessage(hDictionaryListView, LVM_GETHEADER, 0, 0);
	//得到列的数目
	int nCols = SendMessage(hWndListViewHeader, HDM_GETITEMCOUNT, 0, 0);
	//删除所有列	
	nCols--;
	for (; nCols >= 0; nCols--)
		SendMessage(hDictionaryListView, LVM_DELETECOLUMN, nCols, 0);
	
	PAINTSTRUCT ps;
	HDC hdc = BeginPaint(hMainWindow, &ps);
	
	SIZE *textSize = new SIZE();
	
	WCHAR szText[256];// Temporary buffer.
	LVCOLUMN lvc;
	// Initialize the LVCOLUMN structure.
	// The mask specifies that the format, width, text,
	// and subitem members of the structure are valid.
	lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

	//////////////////////////////////////////////////////////
	//添加列名

	//我只能先暂时将列名指定为首行的列名
	auto displayColumnNames = reader->getDisplayColumnNames(listViewRows[0].fileName);
	int diplayColumnCounts = displayColumnNames.size();

	for (int iCol = 0; iCol < diplayColumnCounts; iCol++) 
	{
		lvc.iSubItem = iCol;
		lvc.fmt = LVCFMT_LEFT;		// Left-aligned column.
		
		std::string cn = displayColumnNames[iCol];
		std::wstring columnName = STW_U8(cn);
		LPCWSTR ws = columnName.c_str();
		lvc.pszText = (LPWSTR)ws;  //强制转换
				
		std::wstring content = L"";
		for (auto & si: listViewRows[0].subItems) {
			if (si.columnName == cn) {
				content = STW_U8(si.content);
				break;
			}			
		}

		//统计宽度，决定列所占像素
		unsigned int columnNameWidth = wcslen(columnName.c_str()) * 20;
		unsigned int contentWidth = wcslen(content.c_str()) * 20;	
		unsigned int width = columnNameWidth < contentWidth ? contentWidth : columnNameWidth;
		width = width <= 0 ? 50 : (width > 255 ? 255 : width)  ;
		lvc.cx = width;

		// Insert the columns into the list view.
		if (ListView_InsertColumn(hDictionaryListView, iCol, &lvc) == -1)
			return;

	}

	//for (int iCol = 0; iCol < diplayColumnCounts; iCol++)
	//{
	//	lvc.iSubItem = iCol;
	//	lvc.fmt = LVCFMT_LEFT;		// Left-aligned column.

	//	std::string cn = listViewRows[0].subItems[iCol].columnName;
	//	std::wstring columnName = STW_U8(cn);

	//	LPCWSTR ws = columnName.c_str();
	//	lvc.pszText = (LPWSTR)ws ;  //强制转换
	//	
	//	std::string con = listViewRows[0].subItems[iCol].content;
	//	std::wstring content = STW_U8(con);
	//	
	//	//统计宽度，决定列所占像素

	//	//GetTextExtentPoint32(hdc, lvc.pszText, 256, textSize);
	//	//GetTextExtentPoint32(hdc, lvc.pszText, wcslen(content.c_str()), textSize);
	//	//lvc.cx = textSize->cx;

	//	unsigned int columnNameWidth = wcslen(columnName.c_str()) * 15;
	//	unsigned int contentWidth = wcslen(content.c_str()) * 15;	
	//	unsigned int width = columnNameWidth < contentWidth ? contentWidth : columnNameWidth;
	//	width = width <= 0 ? 50 : (width > 255 ? 255 : width)  ;
	//	lvc.cx = width;

	//	// Insert the columns into the list view.
	//	if (ListView_InsertColumn(hDictionaryListView, iCol, &lvc) == -1)
	//		return;
	//}

	//////////////////////////////////////////////////////////
	//填充数据
	LVITEM vitem;
	vitem.mask = LVIF_TEXT;

	unsigned int rowIndex = 0;

	for (auto& lvr : listViewRows) 
	{
		vitem.iItem = rowIndex;

		for (int iCol = 0; iCol < diplayColumnCounts; iCol++) 
		{
			vitem.iSubItem = iCol;

			for (auto & si : listViewRows[0].subItems) {
				if (si.columnName == displayColumnNames[iCol]) {
					std::wstring wstr = STW_U8(si.content);
					vitem.pszText = (LPWSTR)wstr.c_str();
					
					if (iCol == 0) {
						ListView_InsertItem(hDictionaryListView, &vitem);
					}
					else {
						ListView_SetItem(hDictionaryListView, &vitem);
					}

					break;
				}
			}


		}

		rowIndex++;
	}

	//unsigned int rowIndex = 0;
	//unsigned int columnIndex = 0;
	//for (auto& lvr : listViewRows) 
	//{
	//	vitem.iItem = rowIndex;

	//	for (auto& si : lvr.subItems) {

	//		vitem.iSubItem = columnIndex;
	//		
	//		std::wstring wstr= STW_U8(si.content);
	//		vitem.pszText = (LPWSTR)wstr.c_str();

	//		if (columnIndex == 0) {
	//			ListView_InsertItem(hDictionaryListView, &vitem);
	//		}
	//		else {
	//			ListView_SetItem(hDictionaryListView, &vitem);
	//		}

	//		columnIndex++;
	//	}
	//	
	//	rowIndex++;
	//	columnIndex = 0;
	//}

	//////////////////////////////////////////////////////////
	EndPaint(hMainWindow, &ps);	
}

VOID OpenFiles() 
{
	//调用 shell32.dll api
	//调用浏览文件夹对话框
	TCHAR szPathName[MAX_PATH];
	BROWSEINFO bInfo = { 0 };
	bInfo.hwndOwner = GetForegroundWindow();//父窗口  
	bInfo.lpszTitle = TEXT("选择文件夹");
	bInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI/*包含一个编辑框 用户可以手动填写路径 对话框可以调整大小之类的..*/ |
		BIF_UAHINT/*带TIPS提示*/ | BIF_NONEWFOLDERBUTTON /*不带新建文件夹按钮*/;
	LPITEMIDLIST lpDlist = SHBrowseForFolder(&bInfo);
	if (lpDlist != NULL)//单击了确定按钮  
	{
		//首先清除之前载入的所有信息
		reader->clear();

		////////////////////////////////////////////////////////////////////
		// 获取目录，为了支持中文，需要将获取的字符串的TCHAR编码转为CHAR编码，使用代码页CP_ACP
		SHGetPathFromIDList(lpDlist, szPathName);

		char cPath[256];
		//获取字节长度   
		int iLength = WideCharToMultiByte(CP_ACP, 0, szPathName, -1, NULL, 0, NULL, NULL);
		//将tchar值赋给_char    
		WideCharToMultiByte(CP_ACP, 0, szPathName, -1, cPath, iLength, NULL, NULL);

		////////////////////////////////////////////////////////////////////
		//搜索目录下的xlsx文件，并记录文件名
		std::string searchPath = cPath;
		searchPath = searchPath + "\\*.xlsx";

		std::wstring wpath = szPathName;
		wpath = wpath + L"\\";
		
		_finddata_t fileDir;
		long lfDir;
		if ((lfDir = _findfirst(searchPath.c_str(), &fileDir)) == -1l) {
			MessageBox(hMainWindow, L"No xlsx file is found\n", L"提示", MB_OK);
			return;
		}
		else {
			do {
				//找到了xlsx文件，先将char转换为wchar，再将wchar转换为utf-8编码的char
				TCHAR cache[256];
				iLength = MultiByteToWideChar(CP_ACP, 0, fileDir.name, strlen(fileDir.name) + 1, NULL, 0);
				MultiByteToWideChar(CP_ACP, 0, fileDir.name, strlen(fileDir.name) + 1, cache, iLength);

				std::wstring ws(cache);
				std::string utf8_str = WTS_U8(ws);
				
				reader->addXlsxFile(utf8_str, WTS_U8(wpath));

			} while (_findnext(lfDir, &fileDir) == 0);
		}
		_findclose(lfDir);
		
		////////////////////////////////////////////////////////////////////
		// 激活某些控件
		ActiveControls(TRUE);

		//更新下拉列表内容
		// load the combobox with item list. Send a CB_ADDSTRING message to load each item
		std::vector<std::string>& fileNames = reader->getFileNames();
		TCHAR temp[100];
		int fileCounts = fileNames.size();
		for (int index = 0; index < fileCounts; index++) 
		{
			auto c_str = fileNames[index].c_str();//string---->const char*
			CharToTchar(c_str,temp);
			SendMessage(hFileNameComboBox, (UINT)CB_ADDSTRING, (WPARAM)0, (LPARAM)temp);
		}

		// Send the CB_SETCURSEL message to display an initial item in the selection field  
		SendMessage(hFileNameComboBox, CB_SETCURSEL, (WPARAM)0, (LPARAM)0);

		/////////////////////////////////////////////////
		SetSelectedColumnNamesAndWordColumnName();

	}

}