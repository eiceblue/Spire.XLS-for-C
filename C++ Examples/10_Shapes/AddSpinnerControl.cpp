#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"AddSpinnerControl.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set text for range C11
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C11"))->SetText(L"Value:");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C11"))->GetStyle()->GetFont()->SetIsBold(true);

	//Set value for range B10
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))-> SetNumberValue(0);

	//Add spinner control
	intrusive_ptr<ISpinnerShape> spinner = sheet->GetSpinnerShapes()->AddSpinner(12, 4, 20, 20);
	spinner->SetLinkedCell(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12")));
	spinner->SetMin(0);
	spinner->SetMax(100);
	spinner->SetIncrementalChange(5);
	spinner->SetDisplay3DShading(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
