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
	wstring inputFile = data_path + L"CheckRowOrColumnIsHidden.xlsx";
	wstring outputFile = output_path + L"CheckRowOrColumnIsHidden.txt";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Create a wstring object to store the result.
	wstring* result = new wstring();

	// Set the row and column indices for checking visibility.
	int rowIndex = 2;
	int columnIndex = 2;

	// Check if the specified row is hidden.
	bool rowIsHide = sheet->GetRowIsHide(rowIndex);

	if (rowIsHide)
	{
		// Append a message indicating that the second row is hidden to the result.
		result->append(L"The second row is hidden.\r");
	}
	else
	{
		// Append a message indicating that the second row is not hidden to the result.
		result->append(L"The second row is not hidden.\r");
	}

	// Check if the specified column is hidden.
	bool columnIsHide = sheet->GetColumnIsHide(columnIndex);

	if (columnIsHide)
	{
		// Append a message indicating that the second column is hidden to the result.
		result->append(L"The second column is hidden.\r");
	}
	else
	{
		// Append a message indicating that the second column is not hidden to the result.
		result->append(L"The second column is not hidden.\r");
	}

	// Write the result string to the output file.
	WriteAllText(outputFile, *result);

	// Dispose of the workbook object.
	workbook->Dispose();
}

