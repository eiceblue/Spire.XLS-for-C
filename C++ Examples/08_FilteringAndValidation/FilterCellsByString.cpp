#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"FilterCellsByString.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Get the AutoFiltersCollection from the worksheet.
	intrusive_ptr<AutoFiltersCollection> filters = dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters());

	// Set the range for filtering to column D (D1:D19).
	filters->SetRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1:D19")));

	// Retrieve the first filter column from the AutoFiltersCollection.
	intrusive_ptr<FilterColumn> filtercolumn = dynamic_pointer_cast<FilterColumn>(filters->Get(0));

	// Create a SpireString object for the custom filter criteria ("South*").
	intrusive_ptr<SpireString> cri = new SpireString(L"South*");

	// Apply the custom filter to the filter column using Equal operator and the criteria.
	filters->CustomFilter(filtercolumn, FilterOperatorType::Equal, cri);

	// Apply the filters to filter the data.
	filters->Filter();

	// Save the workbook to the specified output file in Excel 2013 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);

	// Dispose of the workbook object.
	workbook->Dispose();
}

