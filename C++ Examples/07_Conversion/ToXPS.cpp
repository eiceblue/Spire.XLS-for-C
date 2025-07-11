#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToXPS.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToXPS.xps";


		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), Spire::Xls::FileFormat::XPS);
		workbook->Dispose();
}
