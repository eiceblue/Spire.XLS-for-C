#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"AddATotalRowToTable.xlsx";
	wstring outputFile = output_path + L"AddATotalRowToTable_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a table with the data from the specific cell range.
	intrusive_ptr<IListObject> table = sheet->GetListObjects()->Create(L"Table", dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D4")));

	//Display total row.
	table->SetDisplayTotalRow(true);

	//Add a total row.
	intrusive_ptr<Spire::Xls::IList<IListObjectColumn>> list = table->GetColumns();
	list->GetItem(0)->SetTotalsRowLabel(L"Total");
	list->GetItem(1)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
	list->GetItem(2)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);
	list->GetItem(3)->SetTotalsCalculation(ExcelTotalsCalculation::Sum);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

