#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Indentation.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Access the "B5" cell from the worksheet
	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));

	//Add some value to the "B5" cell
	cell->SetText(L"Hello Spire!");

	//Set the indentation level of the text (inside the cell) to 2
	cell->GetStyle()->SetIndentLevel(2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}