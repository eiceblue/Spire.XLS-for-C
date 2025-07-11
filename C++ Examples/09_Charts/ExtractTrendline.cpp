#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample4.xlsx";
	wstring outputFile = output_path + L"ExtractTrendline.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the chart from the first worksheet
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetCharts()->Get(0));

	//Get the trendline of the chart and then extract the equation of the trendline
	intrusive_ptr<IChartTrendLine> trendLine = chart->GetSeries()->Get(1)->GetTrendLines()->GetItem(0);
	wstring formula = trendLine->GetFormula();
	std::wstring* content = new std::wstring();
	content->append(L"The equation is: ");
	content->append(formula);
	std::wfstream out;
	out.open(outputFile, ios::out);

	out << *content << endl;
	out.close();
	workbook->Dispose();
}
