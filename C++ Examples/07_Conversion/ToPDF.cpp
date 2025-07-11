#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToPDF.pdf";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		workbook->GetConverterSetting()->SetSheetFitToPage(true);

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
		workbook->Dispose();
}
