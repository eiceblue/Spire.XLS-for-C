#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring data_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = data_path + L"templateAz2.xlsx";
	std::wstring outputFile = output_path + L"OpenExistingFile.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->LoadFromFile(inputFile.c_str());
	//Add a new sheet, named MySheet
	intrusive_ptr<Worksheet> sheet = workbook->GetWorksheets()->Add(L"MySheet");

	//Get the reference of L"A1" cell from the cells collection of a worksheet
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Hello World");

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();


}

