#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_1.xlsx";
	wstring outputFile = output + L"SetOtherPrintingOptions.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the reference of the PageSetup of the worksheet.
	intrusive_ptr<PageSetup> pageSetup = dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup());

	//Allow to print gridlines.
	pageSetup->SetIsPrintGridlines(true);

	//Allow to print row/column headings.
	pageSetup->SetIsPrintHeadings(true);

	//Allow to print worksheet in black & white mode.
	pageSetup->SetBlackAndWhite(true);

	//Allow to print comments as displayed on worksheet.
	pageSetup->SetPrintComments(PrintCommentType::InPlace);

	//Allow to print worksheet with draft quality.
	pageSetup->SetDraft(true);

	//Allow to print cell errors as N/A.
	pageSetup->SetPrintErrors(PrintErrorsType::NA);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}