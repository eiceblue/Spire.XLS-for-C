
#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"templateAz.xlsx";
	wstring outputFile = output_path + L"GetColorArgbData.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	std::wstring* content = new std::wstring();

	//Get font color
	intrusive_ptr<Spire::Xls::Color> color1 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->GetStyle()->GetFont()->GetColor();

	//Read ARGB data of Color
	int a, r, g, b;
	a = color1->GetA(), r = color1->GetR(), g = color1->GetG(), b = color1->GetB();
	wstring argb = color1->ToString();
	content->append(L"The font color of B2: " + argb);

	intrusive_ptr<Spire::Xls::Color> color2 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->GetStyle()->GetFont()->GetColor();
	a = color2->GetA(), r = color2->GetR(), g = color2->GetG(), b = color2->GetB();
	argb = color2->ToString();
	content->append(L"The font color of B3: " + argb);

	intrusive_ptr<Spire::Xls::Color> color3 = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->GetStyle()->GetFont()->GetColor();
	a = color3->GetA(), r = color3->GetR(), g = color3->GetG(), b = color3->GetB();
	argb = color2->ToString();
	content->append(L"The font color of B4: " + argb);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();

}
