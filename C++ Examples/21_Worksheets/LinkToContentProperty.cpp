#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"AccessDocumentProperties.xlsx";
	std::wstring outputFile = output_path + L"LinkToContentProperty.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Add a custom document property
	workbook->GetCustomDocumentProperties()->Add(L"Test", L"MyNamedRange");
	//Get the added document property
	intrusive_ptr<ICustomDocumentProperties> properties = workbook->GetCustomDocumentProperties();
	intrusive_ptr<IDocumentProperty> property_Renamed = properties->Get(L"Test");
	//Link to content 
	property_Renamed->SetLinkToContent(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}