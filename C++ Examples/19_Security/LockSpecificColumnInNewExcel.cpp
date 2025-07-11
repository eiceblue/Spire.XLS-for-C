#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"LockSpecificColumnInNewExcel.xlsx";
	
	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Create an empty worksheet.
	workbook->CreateEmptySheet();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Loop through all the columns in the worksheet and unlock them.
	for (int i = 0; i < 255; i++)
	{
		sheet->GetRows()->GetItem(i)->GetStyle()->SetLocked(false);
	}

	//Lock the fourth column in the worksheet.
	sheet->GetColumns()->GetItem(3)->SetText(L"Locked");
	sheet->GetColumns()->GetItem(3)->GetStyle()->SetLocked(true);
	//Set the password.
	sheet->XlsWorksheetBase::Protect(L"123");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}