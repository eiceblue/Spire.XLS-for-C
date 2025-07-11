#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"OfficeOpenXMLToExcel.xml";
   	 wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"OfficeOpenXMLToExcel.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Initialize worksheet
		ifstream fs(inputFile.c_str(), ios::in | ios::binary);
		intrusive_ptr<Stream> fileStream = new Stream(fs);
		workbook->LoadFromXml(fileStream);

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
		workbook->Dispose();
}
