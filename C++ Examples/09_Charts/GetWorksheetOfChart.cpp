#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartToImage.xlsx";
	wstring outputFile = output_path + L"GetWorksheetOfChart.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Access the first chart inside this worksheet
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Get its worksheet
	intrusive_ptr<Worksheet> wSheet = chart->GetWorksheet();

	//Create StringBuilder to save 
	std::wstring* content = new std::wstring();

	//Set string format for displaying
	wstring s = L"Sheet Name: ";
	wstring result = s + sheet->GetName() + L"\r\nCharts' sheet Name: " + wSheet->GetName();

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}
