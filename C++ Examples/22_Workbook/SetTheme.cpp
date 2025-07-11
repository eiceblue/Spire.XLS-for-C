#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"SetTheme.xlsx";
	std::wstring outputFile = output_path + L"SetTheme.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> srcWorkbook = new Workbook();

	//Load the Excel document from disk
	srcWorkbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> srcWorksheet = dynamic_pointer_cast<Worksheet>(srcWorkbook->GetWorksheets()->Get(0));

	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->GetWorksheets()->Clear();
	workbook->GetWorksheets()->AddCopy(srcWorksheet);

	//1. Copy the theme of the workbook
	//workbook->CopyTheme(srcWorkbook);

	//2. Set a certain type of color of the default theme in the workbook
	workbook->SetThemeColor(ThemeColorType::Dk1, Spire::Xls::Color::GetSkyBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}