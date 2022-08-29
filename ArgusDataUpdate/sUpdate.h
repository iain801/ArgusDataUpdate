#pragma once
#include "libxl.h"
#include <string>
#include <list>

class sUpdate
{
public:
	sUpdate(std::wstring srcPath, std::wstring destPath, 
		std::wstring sheetLabel);
	~sUpdate();
	
	void CopySheet();
	bool isSheets();

private:
	std::wstring srcPath;
	std::wstring destPath;
	libxl::Book* src = nullptr;
	libxl::Book* dest = nullptr;
	libxl::Sheet* srcSheet = nullptr; 
	libxl::Sheet* destSheet = nullptr;
	
	std::wstring sheetLabel;
	
	void CopyCell(int row, int col);
	void CopyCell(int srcRow, int destRow, int srcCol, int destCol);
	int getSheet(libxl::Book* book, std::wstring label);
};

