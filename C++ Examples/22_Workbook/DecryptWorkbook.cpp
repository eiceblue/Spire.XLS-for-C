#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"DecryptWorkbook.xlsx";
	std::wstring outputFile = output_path + L"DecryptWorkbook.xlsx";

	bool value = Workbook::IsPasswordProtected(inputFile.c_str());

	if (value)
	{
		//Load a file with the password specified
		intrusive_ptr<Workbook> workbook = new Workbook();
		workbook->SetOpenPassword(L"eiceblue");
		workbook->LoadFromFile(inputFile.c_str());

		//Decrypt workbook
		workbook->UnProtect();

		//Save the document
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
		workbook->Dispose();
	}
}