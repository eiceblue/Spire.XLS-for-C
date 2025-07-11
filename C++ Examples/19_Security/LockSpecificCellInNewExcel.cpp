#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"LockSpecificCellInNewExcel.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Create an empty worksheet.
	workbook->CreateEmptySheet();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Loop through all the rows in the worksheet and unlock them.
	for (int i = 0; i < 255; i++)
	{
		sheet->GetRows()->GetItem(i)->GetStyle()->SetLocked(false);
	}

	//Lock specific cell in the worksheet.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Locked");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetStyle()->SetLocked(true);

	//Lock specific cell range in the worksheet.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:E3"))->SetText(L"Locked");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:E3"))->GetStyle()->SetLocked(true);

	//Set the password.
	sheet->XlsWorksheetBase::Protect(L"123");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}