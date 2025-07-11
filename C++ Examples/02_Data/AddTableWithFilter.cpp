#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_4.xlsx";
	wstring outputFile = output_path + L"AddTableWithFilter.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a List Object named in Table.
	sheet->GetListObjects()->Create(L"Table", sheet->GetRange(1, 1, sheet->GetLastRow(), sheet->GetLastColumn()));

	//Set the BuiltInTableStyle for List object.
	intrusive_ptr<IListObjects> ie = sheet->GetListObjects();
	int i = ie->GetCount();
	ie->GetItem(0)->SetBuiltInTableStyle(TableBuiltInStyles::TableStyleLight9);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

