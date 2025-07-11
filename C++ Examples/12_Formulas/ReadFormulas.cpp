#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReadFormulas.xlsx";
	wstring outputFile = output_path + L"ReadFormulas.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	wstring formula = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->GetFormula();
	wstring value = to_wstring(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C14"))->GetFormulaNumberValue());

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << "Formula£º" << formula << "\r\n" << "Value£º" << value << endl;

	ofs.close();
	workbook->Dispose();
}