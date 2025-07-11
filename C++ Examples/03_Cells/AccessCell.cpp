#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"AccessCell.xlsx";
	wstring outputFile = output_path + L"AccessCell_result.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	std::wstring* builder = new std::wstring();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Access cell by its name
	intrusive_ptr<CellRange> range1 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
	wstring s1 = range1->GetText();
	builder->append(L"Value of range1: " + s1 + L"\n");

	//Access cell by index of row and column
	intrusive_ptr<CellRange> range2 = dynamic_pointer_cast<CellRange>(sheet->GetRange(2, 1));
	wstring s2 = range2->GetText();
	builder->append(L"Value of range2: " + s2 + L"\n");

	//Access cell in cell collection
	intrusive_ptr<XlsRange> range3 = sheet->GetCells()->GetItem(2);
	wstring s3 = range3->GetText();
	builder->append(L"Value of range3: " + s3 + L"\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
	workbook->Dispose();

}

