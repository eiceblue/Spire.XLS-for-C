#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReadHyperlinks.xlsx";
	wstring outputFile = output_path + L"ReadHyperlinks.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	wstring address1 = sheet->GetHyperLinks()->Get(0)->GetAddress();
	wstring address2 = sheet->GetHyperLinks()->Get(1)->GetAddress();

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << address1 + L"\r\n" + address2 << endl;
	ofs.close();
	workbook->Dispose();
}