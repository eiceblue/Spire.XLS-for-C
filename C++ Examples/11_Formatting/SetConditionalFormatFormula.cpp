#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetConditionalFormatFormula.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add ConditionalFormat
	intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();

	//Define the range
	xcfs->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5")));

	//Add condition
	intrusive_ptr<IConditionalFormat> format = xcfs->AddCondition();
	format->SetFormatType(ConditionalFormatType::CellValue);

	//If greater than 1000
	format->SetFirstFormula(L"1000");
	format->SetOperator(ComparisonOperatorType::Greater);
	format->SetBackColor(Spire::Xls::Color::GetOrange());

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetNumberValue(40);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(500);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(300);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(400);

	//Set a SUM formula for B5
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetFormula(L"=SUM(B1:B4)");

	//Add text
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetText(L"If Sum of B1:B4 is greater than 1000, B5 will have orange background.");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}