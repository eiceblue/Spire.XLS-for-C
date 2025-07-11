#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ReadImages.xlsx";
	std::wstring outputFile = output_path + L"DetectEmptyWorksheet.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> worksheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Detect the first worksheet is empty or not
	bool detect1 = worksheet1->GetIsEmpty();

	//Get the second worksheet
	intrusive_ptr<Worksheet> worksheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

	//Detect the second worksheet is empty or not
	bool detect2 = worksheet2->GetIsEmpty();

	//Create StringBuilder to save 
	wstring* content = new wstring();

	//Set string format for displaying
	wstring result = L"The first worksheet is empty or not: " + to_wstring(detect1) + L"\r\nThe second worksheet is empty or not: " + to_wstring(detect2);

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}