#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AutofilterSample.xlsx";
   	wstring output_path = OUTPUTPATH;
   	wstring outputFile = output_path + L"ToCSVWithFilteredValue.csv";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Convert to CSV file with filtered value
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->SaveToFile(outputFile.c_str(), L";", false);
	workbook->Dispose();
}
