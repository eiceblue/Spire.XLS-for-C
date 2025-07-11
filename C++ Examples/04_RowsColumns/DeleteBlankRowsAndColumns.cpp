#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Template_Xls_2.xlsx";
	wstring outputFile = outputFolder + L"DeleteBlankRowsAndColumns_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Delete blank rows from the worksheet.
	for (int i = sheet->GetRows()->GetCount() - 1; i >= 0; i--)
	{
		if (sheet->GetRows()->GetItem(i)->GetIsBlank())
		{
			sheet->DeleteRow(i + 1);
		}
	}

	//Delete blank columns from the worksheet.
	for (int j = sheet->GetColumns()->GetCount() - 1; j >= 0; j--)
	{
		if (sheet->GetColumns()->GetItem(j)->GetIsBlank())
		{
			sheet->DeleteColumn(j + 1);
		}
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
