#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImportDataFromArrayList_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Create an empty worksheet
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create an ArrayList object
	vector<LPCWSTR_S> list;

	//Add strings in list
	list.push_back(L"Spire.Doc for C++");
	list.push_back(L"Spire.XLS for C++");
	list.push_back(L"Spire.PDF for C++");
	list.push_back(L"Spire.Presentation for C++");

	//Insert arrary list in worksheet 
	sheet->InsertArray(list, 1, 1, true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

