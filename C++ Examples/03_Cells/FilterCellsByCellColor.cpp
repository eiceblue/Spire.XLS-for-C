#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_3.xlsx";
	wstring outputFile = output_path + L"FilterCellsByCellColor_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create an auto filter in the sheet and specify the range to be filterd
	sheet->GetAutoFilters()->SetRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"G1:G19")));

	//Get the coloumn to be filterd
	intrusive_ptr<FilterColumn> filtercolumn = sheet->GetAutoFilters()->Get(0);

	//Add a color filter to filter the column based on cell color
	(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->AddFillColorFilter(filtercolumn, Spire::Xls::Color::GetRed());

	//Filter the data.
	(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Filter();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
