#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"AllNamedRanges.xlsx";
	wstring outputFile = output_path + L"GetAllNamedRange.txt";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	wstring* content = new wstring();

	//Get all named range
	intrusive_ptr<INameRanges> ranges = workbook->GetNameRanges();
	for (int i = 0; i < ranges->GetCount(); i++)
	{
		intrusive_ptr<INamedRange> nameRange = ranges->Get(i);
		content->append(nameRange->GetName());
		content->append(L"\r\n");
	}

	//Save to file.
	std::wfstream out;
	out.open(outputFile, ios::out);

	out << *content << endl;
	out.close();
	workbook->Dispose();
}