#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring picPath = input_path + L"SpireXls.png";
	wstring outputFile = output_path + L"AddImageHyperlink.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add the description text
	sheet->GetColumns()->GetItem(0)->SetColumnWidth(22);
	sheet->GetRange(L"A1")->SetText(L"Image Hyperlink");
	sheet->GetRange(L"A1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Top);

	//Insert an image to a specific cell
	intrusive_ptr<ExcelPicture> picture = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(2, 1, picPath.c_str()));
	//Add a hyperlink to the image
	picture->SetHyperLink(L"https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}