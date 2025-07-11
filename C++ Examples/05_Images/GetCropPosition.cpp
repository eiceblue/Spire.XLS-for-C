#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"CropPosition.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the image from the first sheet
	intrusive_ptr<XlsBitmapShape> picture = sheet->GetPictures()->Get(0);

	//Get the cropped position
	int left = picture->GetLeft();
	int top = picture->GetTop();
	int width = picture->GetWidth();
	int height = picture->GetHeight();

	//Create StringBuilder to save 
	std::wstring* content = new std::wstring();

	//Set string format for displaying
	wstring displayString = L"Crop position: Left " + std::to_wstring(left) + L"\r\nCrop position: Top " + std::to_wstring(top) + L"\r\nCrop position: Width " + std::to_wstring(width) + L"\r\nCrop position: Height " + std::to_wstring(height);

	//Add result string to StringBuilder
	content->append(displayString);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}
