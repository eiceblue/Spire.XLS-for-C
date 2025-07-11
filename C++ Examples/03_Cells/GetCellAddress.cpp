#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"DataSorting.xls";
	wstring outputFile = output_path + L"GetCellAddress_result.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	std::wstring* builder = new std::wstring();

	//Get a cell range
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B5"));

	//Get address of range
	wstring address = range->GetRangeAddressLocal();
	builder->append(L"Address of range: " + address);

	//Get the cell count of range
	int count = range->GetCellsCount();
	builder->append(L"Cell count of range: " + std::to_wstring(count));

	//Get the address of the entire column of range
	wstring entireColAddress = range->GetEntireColumn()->GetRangeAddressLocal();
	builder->append(L"Address of entire column of the range: " + entireColAddress);

	//Get the address of the entire row of range
	wstring entireRowAddress = range->GetEntireColumn()->GetRangeAddressLocal();
	builder->append(L"Address of entire row of the range " + entireRowAddress);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << builder << endl;
	ofs.close();
	workbook->Dispose();
}

