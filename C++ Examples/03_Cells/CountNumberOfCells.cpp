#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring inputFile = data_path + L"Template_Xls_4.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CountNumberOfCells.txt";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	std::wstring* content = new std::wstring();

	//Get the number of cells.
	content->append(L"Number of Cells: " + std::to_wstring(sheet->GetCells()->GetCount()));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

