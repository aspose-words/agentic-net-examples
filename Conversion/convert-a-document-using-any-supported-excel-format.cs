using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input Word document path (can be any supported format).
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Output Excel file path (XLSX format).
        string outputPath = @"C:\Docs\ConvertedDocument.xlsx";

        // Load the document using the Document constructor (lifecycle rule).
        Document doc = new Document(inputPath);

        // Configure XlsxSaveOptions (specific to Excel output).
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Ensure the save format is set to Xlsx (required by the options).
            SaveFormat = SaveFormat.Xlsx,

            // Example: save each section to a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document to the Excel format using the Save method (lifecycle rule).
        doc.Save(outputPath, xlsxOptions);
    }
}
