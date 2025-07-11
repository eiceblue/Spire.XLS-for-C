#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"DataValidation_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->SetText(L"Please enter time between 09:00 and 18:00:");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C12"))->AutoFitColumns();

		//Set Time data validation for cell L"D12"
		intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D12"));
		range->GetDataValidation()->SetAllowType(CellDataType::Time);
		range->GetDataValidation()->SetCompareOperator(ValidationComparisonOperator::Between);

		range->GetDataValidation()->SetFormula1(L"09:00");
		range->GetDataValidation()->SetFormula2(L"18:00");

		range->GetDataValidation()->SetAlertStyle(AlertStyleType::Info);
		range->GetDataValidation()->SetShowError(true);
		range->GetDataValidation()->SetErrorTitle(L"Time Error");
		range->GetDataValidation()->SetErrorMessage(L"Please enter a valid time");
		range->GetDataValidation()->SetInputMessage(L"Time Validation Type");
		range->GetDataValidation()->SetIgnoreBlank(true);
		range->GetDataValidation()->SetShowInput(true);

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
