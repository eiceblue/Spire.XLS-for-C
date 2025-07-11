#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_3.xlsx";
	wstring outputFile = output_path + L"ExpandAndCollapseGroups_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Expand the grouped rows with ExpandCollapseFlags set to expand parent
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A16:G19"))->ExpandGroup(GroupByType::ByRows, ExpandCollapseFlags::ExpandParent);

	//Collapse the grouped rows
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:G12"))->CollapseGroup(GroupByType::ByRows);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

