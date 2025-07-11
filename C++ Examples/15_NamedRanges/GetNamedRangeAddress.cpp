#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"AllNamedRanges.xlsx";
	wstring outputFile = output_path + L"GetNamedRangeAddress.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	wstring* content = new wstring();

	//Get specific named range by index
	intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Get(0);

	//Get the address of the named range
	wstring address = NamedRange->GetRefersToRange()->GetRangeAddress();
	content->append(L"The address of the named range ");
	content->append(NamedRange->GetName());
	content->append(L" is ");
	content->append(address);

	//Save to file.
	ofs.open(outputFile, ios::out);
	wstring result(content->begin(), content->end());
	ofs << result << endl;
	ofs.close();
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}