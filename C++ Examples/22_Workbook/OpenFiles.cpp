#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	std::wstring inputFile_97 = input_path + L"ExcelSample97_N.xls";
	std::wstring inputFile_xml = input_path + L"OfficeOpenXML_N.xml";
	std::wstring inputFile_csv = input_path + L"CSVSample_N.csv";
	std::wstring outputFile = output_path + L"OpenFiles.txt";
	wfstream ofs;


	//Create wstring builder
	wstring* content = new wstring();

	//1. Load file by file path
	//Create a workbook
	intrusive_ptr<Workbook> workbook1 = new Workbook();
	//Load the document from disk
	workbook1->LoadFromFile(inputFile.c_str());
	content->append(L"Workbook opened using file path successfully!");

	//2. Load file by file stream 
	ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
	intrusive_ptr<Stream> stream = new Stream(inputf);
	//Create a workbook
	intrusive_ptr<Workbook> workbook2 = new Workbook();
	//Load the document from disk
	workbook2->LoadFromStream(stream);
	content->append(L"\r\nWorkbook opened using file stream successfully!");
	//delete stream;

	//3. Open Microsoft Excel 97 - 2003 file
	intrusive_ptr<Workbook> wbExcel97 = new Workbook();
	wbExcel97->LoadFromFile(inputFile_97.c_str(), ExcelVersion::Version97to2003);
	content->append(L"\r\nMicrosoft Excel 97 - 2003 workbook opened successfully!");

	//4. Open xml file
	intrusive_ptr<Workbook> wbXML = new Workbook();
	wbXML->LoadFromXml(inputFile_xml.c_str());
	content->append(L"\r\nXML file opened successfully!");

	//5. Open csv file
	intrusive_ptr<Workbook> wbCSV = new Workbook();
	wbCSV->LoadFromFile(inputFile_csv.c_str(), L",", 1, 1);
	content->append(L"\r\nCSV file opened successfully!");

	//Save to file.
	std::wfstream out;
	out.open(outputFile, ios::out);

	out << *content << endl;
	out.close();
}