#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"VerifyDataByValidation.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Cell B4 has the Decimal Validation
	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"));

	//Get the valditation of this cell
	intrusive_ptr<Validation> validation = cell->GetDataValidation();

	//Get the specified data range
	double minimum = std::stod(validation->GetFormula1());
	double maximum = std::stod(validation->GetFormula2());

	//Create StringBuilder to save 
	std::wstring* content = new std::wstring();

	//Set different numbers for the cell
	for (int i = 5; i < 100; i = i + 40)
	{
		cell->SetNumberValue(i);
		std::wstring result = L"";
		//Verify 
		if (cell->GetNumberValue() < minimum || cell->GetNumberValue() > maximum)
		{
			//Set string format for displaying
			result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: false";
		}
		else
		{
			//Set string format for displaying
			result = L"Is input " + std::to_wstring(i) + L" a valid value for this Cell: true";
		}
		//Add result string to StringBuilder
		content->append(result);
	}


	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}
