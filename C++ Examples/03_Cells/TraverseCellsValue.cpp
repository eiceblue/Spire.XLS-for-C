#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CellValues.xlsx";
	wstring outputFile = outputFolder + L"TraverseCellsValue_out.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the cell range collection 
	intrusive_ptr<Spire::Xls::IList<XlsRange>> cellRangeCollection = sheet->GetCells();

	//Create StringBuilder to save 
	std::wstring* content = new std::wstring();
	content->append(L"Values of the first sheet:");

	//Traverse cells value
	//for (auto cellRange : cellRangeCollection)
	//{
	for (int i = 0; i < cellRangeCollection->GetCount(); i++)
	{
		intrusive_ptr<XlsRange> cr = cellRangeCollection->GetItem(i);
		//Set string format for displaying
		wstring cell = cr->GetRangeAddress();
		wstring result = L"Cell: " + cell + L"   Value: " + cr->GetValue();

		//Add result string to StringBuilder
		content->append(result);
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}
