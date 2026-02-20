using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Example.docx";

        // Path where the resulting Excel file will be saved.
        string outputPath = @"C:\Docs\Example.xlsx";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Configure save options for XLSX format.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Create a separate worksheet for each section of the document.
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document as an Excel workbook.
        doc.Save(outputPath, saveOptions);
    }
}
