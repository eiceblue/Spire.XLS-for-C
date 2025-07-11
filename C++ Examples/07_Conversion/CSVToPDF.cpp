#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CSVSample.csv";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"CSVToPDF.pdf";

		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str(), L",");

		//Set the SheetFitToPage property as true
		workbook->GetConverterSetting()->SetSheetFitToPage(true);

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Autofit a column if the characters in the column exceed column width
		for (int i = 1; i < sheet->GetColumns()->GetCount(); i++)
		{
			sheet->AutoFitColumn(i);
		}

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
		workbook->Dispose();
}
