#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"CopyVisibleSheets.xlsx";
	std::wstring outputFile = output_path + L"CopyVisibleSheets.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Create a new workbook
	intrusive_ptr<Workbook> workbookNew = new Workbook();
	workbookNew->SetVersion(ExcelVersion::Version2013);
	workbookNew->GetWorksheets()->Clear();

	//Loop through the worksheets
	for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
	{
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
		//Judge if the worksheet is visible
		if (sheet->GetVisibility() == WorksheetVisibility::Visible)
		{
			//Copy the sheet to new workbook
			wstring name = sheet->GetName();
			workbookNew->GetWorksheets()->AddCopy(sheet);
		}
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}