#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConversionSample1.xlsx";
   	 wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"SelectedRangeToPDF.pdf";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Add a new sheet to workbook
		workbook->GetWorksheets()->Add(L"newsheet");
		//Copy your area to new sheet.
		dynamic_pointer_cast<CellRange>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRange(L"A9:E15"))->Copy(dynamic_pointer_cast<CellRange>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetRange(L"A9:E15")), false, true);
		//Auto fit column width
		dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetRange(L"A9:E15")->AutoFitColumns();

		//Save to file.
		dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->SaveToPdf(outputFile.c_str());
		workbook->Dispose();
}
