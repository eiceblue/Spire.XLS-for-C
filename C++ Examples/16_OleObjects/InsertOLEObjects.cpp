#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"InsertOLEObjects.xls";
	wstring outputFile = output_path + L"InsertOLEObjects.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	sheet->GetRange(L"A1")->SetText(L"Here is an OLE Object.");
	//insert OLE object
	intrusive_ptr<Workbook> book = new Workbook();
	book->LoadFromFile(inputFile.c_str());
	book->GetWorksheets()->Get(0)->GetPageSetup()->SetLeftMargin(0);
	book->GetWorksheets()->Get(0)->GetPageSetup()->SetRightMargin(0);
	book->GetWorksheets()->Get(0)->GetPageSetup()->SetTopMargin(0);
	book->GetWorksheets()->Get(0)->GetPageSetup()->SetBottomMargin(0);
	intrusive_ptr<Stream> image = book->GetWorksheets()->Get(0)->ToImage(1, 1, 19, 5);
	intrusive_ptr<Spire::Xls::IOleObject> oleObject = sheet->GetOleObjects()->Add(inputFile.c_str(), image, OleLinkType::Embed);

	oleObject->SetLocation(sheet->GetRange(L"B4"));
	oleObject->SetObjectType(OleObjectType::ExcelWorksheet);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}