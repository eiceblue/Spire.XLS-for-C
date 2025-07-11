#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ProtectedWorksheet.xlsx";
	std::wstring outputFile = output_path + L"VerifyProtectedWorksheet.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Verify the first worksheet 
	bool detect = sheet->GetIsPasswordProtected();

	//Create StringBuilder to save 
	wstring* content = new wstring();

	//Set string format for displaying
	wstring result = L"The first worksheet is password protected or not: " + to_wstring(detect);

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}