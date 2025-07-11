#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"AllNamedRanges.xlsx";
	wstring outputFile = output_path + L"GetSpecificNamedRange.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	wstring* content = new wstring();

	//Get specific named range by index
	wstring name1 = workbook->GetNameRanges()->Get(1)->GetName();
	content->append(L"Get the specific named range " + name1 + L" by index" + L"\r\n");


	//Get specific named range by name
	wstring name2 = workbook->GetNameRanges()->Get(L"NameRange3")->GetName();
	content->append(L"Get the specific named range " + name2 + L" by name" + L"\r\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	wstring result(content->begin(), content->end());
	ofs << result << endl;
	ofs.close();
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}
