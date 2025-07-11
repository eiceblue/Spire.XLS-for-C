#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = data_path + L"AddCustomProperties.xlsx";
	std::wstring outputFile = output_path + L"AddCustomProperties.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Add a custom property ro make the document as final
	workbook->GetCustomDocumentProperties()->Add(L"_MarkAsFinal", true);

	//Add other custom properties to the workbook
	workbook->GetCustomDocumentProperties()->Add(L"The Editor", L"E-iceblue");
	workbook->GetCustomDocumentProperties()->Add(L"Phone number", 81705109);
	workbook->GetCustomDocumentProperties()->Add(L"Revision number", 7.12);
	tm t;
	t.tm_year = 2021 - 1900;
	t.tm_mon = 1 - 1;
	t.tm_mday = 8;
	t.tm_hour = 8;//beijing zone must +8
	t.tm_min = 0;
	t.tm_sec = 0;
	intrusive_ptr<Spire::Xls::DateTime> dt = new Spire::Xls::DateTime(2021, 1, 8, 0, 0, 0);
	workbook->GetCustomDocumentProperties()->Add(L"Revision date", dt);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}