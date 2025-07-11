#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplyGradientFillEffects.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->SetVersion(ExcelVersion::Version2010);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get "B5" cell
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));
	//Set row height and column width
	range->SetRowHeight(50);
	range->SetColumnWidth(30);
	range->SetText(L"Hello");

	//Set alignment style
	range->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	//Set gradient filling effects
	range->GetStyle()->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);
	range->GetStyle()->GetInterior()->GetGradient()->SetForeColor(Spire::Xls::Color::FromArgb(255, 255, 255));
	range->GetStyle()->GetInterior()->GetGradient()->SetBackColor(Spire::Xls::Color::FromArgb(79, 129, 189));
	range->GetStyle()->GetInterior()->GetGradient()->TwoColorGradient(GradientStyleType::Horizontal, GradientVariantsType::ShadingVariants1);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}
