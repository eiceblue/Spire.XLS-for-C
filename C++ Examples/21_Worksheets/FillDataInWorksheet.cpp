#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"FillDataInWorksheet.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Fill data
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Month");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetText(L"January");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetText(L"February");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetText(L"March");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetText(L"April");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetText(L"Payments");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(251);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(515);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(454);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(874);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetText(L"Sample");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetText(L"Sample1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetText(L"Sample2");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetText(L"Sample3");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetText(L"Sample4");

	//Set width for the second column
	sheet->SetColumnWidth(2, 10);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}