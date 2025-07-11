#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"SampeB_4.xlsx";
	wstring outputFile = output_path + L"GetCategoryLabels.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the chart
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Get the cell range of the category labels
	std::wstring* content = new std::wstring();
	intrusive_ptr<IXLSRange> cr = dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetCategoryLabels();
	for (int i = 0; i < cr->GetCells()->GetCount(); i++)
	{
		auto cell = cr->GetCells()->GetItem(i);
		wstring value = cell->GetValue();
		content->append(value + L"\r\n");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}
