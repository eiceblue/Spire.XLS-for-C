#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ListDataValidation_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Set text for cells 
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"Beijing");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8"))->SetText(L"New York");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9"))->SetText(L"Denver");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10"))->SetText(L"Paris");

		//Set data validation for cell
		intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"));
		range->GetDataValidation()->SetShowError(true);
		range->GetDataValidation()->SetAlertStyle(AlertStyleType::Stop);
		range->GetDataValidation()->SetErrorTitle(L"Error");
		range->GetDataValidation()->SetErrorMessage(L"Please select a city from the list");
		range->GetDataValidation()->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:A10")));

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
