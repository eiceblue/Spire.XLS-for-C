#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_3.xlsx";
	wstring outputFile = output_path + L"RetrieveAndExtractData_result.xlsx";

	// Create a new workbook instance and get the first worksheet.
	intrusive_ptr<Workbook> newBook = new Workbook();
	intrusive_ptr<Worksheet> newSheet = dynamic_pointer_cast<Worksheet>(newBook->GetWorksheets()->Get(0));

	//Create a new workbook instance and load the sample Excel file.
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet.
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Retrieve data and extract it to the first worksheet of the new excel workbook.
	int k = 1;
	int columnCount = sheet->GetColumns()->GetCount();
	//for (intrusive_ptr<XlsRange> range : sheet->GetColumns()[0]->GetCells())
	//{
	for (int i = 0; i < sheet->GetColumns()->GetItem(0)->GetCells()->GetCount(); i++)
	{
		intrusive_ptr<XlsRange> range = sheet->GetColumns()->GetItem(0)->GetCells()->GetItem(i);
		if (wcscmp(range->GetText(), L"teacher") == 0)
		{
			int x = range->GetRow();
			intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(range->GetRow(), 1, range->GetRow(), columnCount));
			intrusive_ptr<CellRange> destRange = dynamic_pointer_cast<CellRange>(newSheet->GetRange(k + 1, 1, k + 1, columnCount));
			sheet->Copy(sourceRange, destRange, true);
			k++;
		}
	}

	//Save to file.
	newBook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	newBook->Dispose();
}

