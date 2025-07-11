#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_5.xlsx";
	wstring outputFile = output_path + L"ExtractTextImageFromShape.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Extract text from the first shape and save to a txt file.
	intrusive_ptr<IPrstGeomShape> shape1 = sheet->GetPrstGeomShapes()->Get(2);
	wstring s = shape1->GetText();
	wstring* content = new wstring();
	content->append(L"The text in the third shape is: " + s);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}