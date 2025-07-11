#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"templateAz.xlsx";
	std::wstring outputFile = output_path + L"GetListOfFontsUsed.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	std::vector<intrusive_ptr<ExcelFont>> fonts;

	//Loop all sheets of workbook
	for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
	{
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
		for (int r = 0; r < sheet->GetRows()->GetCount(); r++)
		{
			for (int c = 0; c < sheet->GetRows()->GetItem(r)->GetCells()->GetCount(); c++)
			{
				//Get the font of cell and add it to list
				fonts.push_back(dynamic_pointer_cast<ExcelFont>(sheet->GetRows()->GetItem(r)->GetCells()->GetItem(c)->GetStyle()->GetFont()));
			}
		}
	}
	wstring* strB = new wstring();

	for (auto font : fonts)
	{
		strB->append(L"FontName:");
		strB->append(font->GetFontName());
		strB->append(L"; FontSize:{1}");
		strB->append(to_wstring(font->GetSize()));
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *strB << endl;
	ofs.close();
	workbook->Dispose();
}