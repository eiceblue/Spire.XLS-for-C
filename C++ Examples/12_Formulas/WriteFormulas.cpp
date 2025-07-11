#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"WriteFormulas.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	int currentRow = 1;
	wstring currentFormula = L"";

	sheet->SetColumnWidth(1, 32);
	sheet->SetColumnWidth(2, 16);
	sheet->SetColumnWidth(3, 16);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 1))->SetValue(L"Examples of formulas :");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetValue(L"Test data:");

	intrusive_ptr<CellRange> range(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1")));
	range->GetStyle()->GetFont()->SetIsBold(true);
	range->GetStyle()->SetFillPattern(ExcelPatternType::Solid);
	range->GetStyle()->SetKnownColor(ExcelColors::LightGreen1);
	range->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Medium);

	//test data
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetNumberValue(7.3);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 3))->SetNumberValue(5);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 4))->SetNumberValue(8.2);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 5))->SetNumberValue(4);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 6))->SetNumberValue(3);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 7))->SetNumberValue(11.3);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetValue(L"Formulas");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetValue(L"Results");
	range = dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1, currentRow, 2));
	//range.Value(L"Formulas")
	range->GetStyle()->GetFont()->SetIsBold(true);
	range->GetStyle()->SetKnownColor(ExcelColors::LightGreen1);
	range->GetStyle()->SetFillPattern(ExcelPatternType::Solid);
	range->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Medium);
	//str.
	currentFormula = (L"=\"hello\"");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(L"=\"hello\"");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());
	wstring s(L"≤‚ ‘");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 3))->SetFormula((L"=\"" + s + L"\"").c_str());

	//int.
	currentFormula = (L"=300");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	// float
	currentFormula = (L"=3389.639421");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	//bool.
	currentFormula = (L"=false");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=1+2+3+4+5-6-7+8-9");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=33*3/4-2+10");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());


	// sheet reference
	currentFormula = (L"=Sheet1!$B$3");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	// sheet area reference
	currentFormula = (L"=AVERAGE(Sheet1!$D$3:G$3)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());

	// Functions
	currentFormula = (L"=Count(3,5,8,10,2,34)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());


	currentFormula = (L"=NOW()");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->SetFormula(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 2))->GetStyle()->SetNumberFormat(L"yyyy-MM-DD");

	currentFormula = (L"=SECOND(11)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(++currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MINUTE(12)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MONTH(9)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=DAY(10)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=TIME(4,5,7)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=DATE(6,4,2)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=RAND()");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=HOUR(12)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MOD(5,3)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=WEEKDAY(3)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=YEAR(23)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=NOT(true)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=OR(true)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=AND(TRUE)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=VALUE(30)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=LEN(\"world\")");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MID(\"world\",4,2)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=ROUND(7,3)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=SIGN(4)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=INT(200)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=ABS(-1.21)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=LN(15)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=EXP(20)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=SQRT(40)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=PI()");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=COS(9)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=SIN(45)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MAX(10,30)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=MIN(5,7)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=AVERAGE(12,45)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=SUM(18,29)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=IF(4,2,2)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	currentFormula = (L"=SUBTOTAL(3,Sheet1!B2:E3)");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow, 1))->SetText(currentFormula.c_str());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(currentRow++, 2))->SetFormula(currentFormula.c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}