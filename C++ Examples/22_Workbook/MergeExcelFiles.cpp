#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"MergeExcelFiles.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	std::vector<std::wstring> files;
	files.push_back(input_path + L"MergeExcelFiles-1.xlsx");
	files.push_back(input_path + L"MergeExcelFiles-2.xls");
	files.push_back(input_path + L"MergeExcelFiles-3.xlsx");

	intrusive_ptr<Workbook> newbook = new Workbook();
	newbook->SetVersion(ExcelVersion::Version2013);
	//Clear all worksheets
	newbook->GetWorksheets()->Clear();

	//Create a workbook
	intrusive_ptr<Workbook> tempbook = new Workbook();

	for (auto file : files)
	{
		//Load the file
		tempbook->LoadFromFile(file.c_str());
		for (int i = 0; i < tempbook->GetWorksheets()->GetCount(); i++)
		{
			intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(tempbook->GetWorksheets()->Get(i));
			//Copy every sheet in a workbook
			(dynamic_pointer_cast<XlsWorksheetsCollection>(newbook->GetWorksheets()))->AddCopy(sheet, WorksheetCopyType::CopyAll);
		}
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}