#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	std::wstring outputFile_xlsx = output_path + L"SaveFiles_ToXlsx.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Save to file.
	workbook->SaveToFile(outputFile_xlsx.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}