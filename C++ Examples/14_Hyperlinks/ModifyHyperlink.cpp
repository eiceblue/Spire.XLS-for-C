#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ModifyHyperlink.xlsx";
	wstring outputFile = output_path + L"ModifyHyperlink.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Change the values of TextToDisplay and Address property 
	intrusive_ptr<IHyperLinks> links = sheet->GetHyperLinks();
	links->Get(0)->SetTextToDisplay(L"Product livedemo");
	links->Get(0)->SetAddress(L"https://www.e-iceblue.com/LiveDemo.html");
	
	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}