#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring inputFile = fn + L"ReadImages.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AutoFitBasedOnCellValue_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set value for B8
	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"));
	cell->SetText(L"Welcome to Spire.XLs!");

	//Set the cell style
	intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(cell->GetStyle());
	style->GetFont()->SetSize(16);
	style->GetFont()->SetIsBold(true);

	//Autofit column width and row height based on cell value
	cell->AutoFitColumns();
	cell->AutoFitRows();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

