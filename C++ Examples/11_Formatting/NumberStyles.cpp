#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"NumberStyles.xlsx";
	wstring outputFile = output_path + L"NumberStyles.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Input a number value for the specified cell and set the number format
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->SetText(L"NUMBER FORMATTING");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->GetStyle()->GetFont()->SetIsBold(true);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetText(L"0");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C13"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C13"))->SetNumberFormat(L"0");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))->SetText(L"0.00");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->SetNumberFormat(L"0.00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B15"))->SetText(L"#,##0.00");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C15"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C15"))->SetNumberFormat(L"#,##0.00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B16"))->SetText(L"$#,##0.00");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C16"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C16"))->SetNumberFormat(L"$#,##0.00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B17"))->SetText(L"0;[Red]-0");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C17"))->SetNumberValue(-1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C17"))->SetNumberFormat(L"0;[Red]-0");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B18"))->SetText(L"0.00;[Red]-0.00");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C18"))->SetNumberValue(-1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C18"))->SetNumberFormat(L"0.00;[Red]-0.00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B19"))->SetText(L"#,##0;[Red]-#,##0");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C19"))->SetNumberValue(-1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C19"))->SetNumberFormat(L"#,##0;[Red]-#,##0");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B20"))->SetText(L"#,##0.00;[Red]-#,##0.000");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C20"))->SetNumberValue(-1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C20"))->SetNumberFormat(L"#,##0.00;[Red]-#,##0.00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B21"))->SetText(L"0.00E+00");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C21"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C21"))->SetNumberFormat(L"0.00E+00");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B22"))->SetText(L"0.00%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C22"))->SetNumberValue(1234.5678);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C22"))->SetNumberFormat(L"0.00%");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13:B22"))->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);

	//AutoFit Column
	sheet->AutoFitColumn(2);
	sheet->AutoFitColumn(3);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}