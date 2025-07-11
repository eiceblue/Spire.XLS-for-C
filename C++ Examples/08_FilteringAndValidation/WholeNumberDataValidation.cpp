#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"WholeNumberDataValidation.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->SetText(L"Please enter number between 10 and 100:");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->AutoFitColumns();

		//Set Whole Number data validation for cell L"D12"
		intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D12"));
		range->GetDataValidation()->SetAllowType(CellDataType::Integer);
		range->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);

		range->GetDataValidation()->SetFormula1(L"10");
		range->GetDataValidation()->SetFormula2(L"100");

		range->GetDataValidation()->SetAlertStyle(AlertStyleType::Info);
		range->GetDataValidation()->SetShowError(true);
		range->GetDataValidation()->SetErrorTitle(L"Error");
		range->GetDataValidation()->SetErrorMessage(L"Please enter a valid number");
		range->GetDataValidation()->SetInputMessage(L"Whole Number Validation Type");
		range->GetDataValidation()->SetIgnoreBlank(true);
		range->GetDataValidation()->SetShowInput(true);

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
