#include "pch.h"
using namespace Spire::Xls;

// Convert a wide string (wstring) to a regular string (string)
string wstring2string(const std::wstring& wstr)
{
    std::string result;
    result.reserve(wstr.size());
    for (size_t i = 0; i < wstr.size(); ++i)
    {
        result += static_cast<char>(wstr[i] & 0xFF);
    }
    return result;
}

// Write content to a file specified by the outputFilePath
void WriteAllText(wstring outputFilePath, wstring contentStr)
{
    std::wofstream outputFile(wstring2string(outputFilePath), std::ios::out | std::ios::binary);

    outputFile.imbue(std::locale(""));

    outputFile << contentStr;

    outputFile.close();
}

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CheckAutoFitRowsAndColumns.xlsx";
	wstring outputFile = output_path + L"CheckAutoFitRowOrColumn.txt";

	// Create a Workbook object
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Create a wstring object to store the result
	wstring* result = new wstring();

	// Load the input file into the workbook
	workbook->LoadFromFile(inputFile.c_str());

	// Check if the second row is auto fit
	bool isRowAutofit = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRowIsAutoFit(2);
	if (isRowAutofit)
	{
		result->append(L"The second row is auto fit row height.\r");
	}
	else
	{
		result->append(L"The second row is not auto fit row height.\r");
	}

	// Check if the second column is auto fit
	bool isColAutofit = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetColumnIsAutoFit(2);
	if (isColAutofit)
	{
		result->append(L"The second column is auto fit column width.\r");
	}
	else
	{
		result->append(L"The second column is not auto fit column width.\r");
	}

	// Write the result to the output file
	WriteAllText(outputFile, *result);

	// Dispose the workbook to release resources
	workbook->Dispose();
}

