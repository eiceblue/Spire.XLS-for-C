#include "pch.h"
using namespace Spire::Xls;

int main() {
                wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SparkLine.xlsx";
	wstring outputFile = output_path + L"SparkLine.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add sparkline
	intrusive_ptr<SparklineGroup> sparklineGroup = sheet->GetSparklineGroups()->AddGroup(SparklineType::Line);
	intrusive_ptr<SparklineCollection> sparklines = sparklineGroup->Add();
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D2")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:D3")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E3")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:D4")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E4")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:D5")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E5")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:D6")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E6")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:D7")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E7")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:D8")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E8")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9:D9")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E9")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:D10")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10")));
	sparklines->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A11:D11")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E11")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
