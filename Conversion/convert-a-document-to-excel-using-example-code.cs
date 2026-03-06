using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentToExcelConverter
{
    static void Main()
    {
        // Path to the source Word document.
        string inputFile = @"C:\Docs\SampleDocument.docx";

        // Path where the resulting Excel file will be saved.
        string outputFile = @"C:\Docs\SampleDocument.xlsx";

        // Load the Word document from the file system.
        Document doc = new Document(inputFile);

        // Option 1: Directly save using the SaveFormat enumeration.
        // This uses the Document.Save(string, SaveFormat) overload.
        doc.Save(outputFile, SaveFormat.Xlsx);

        // Option 2: Use XlsxSaveOptions for more control (e.g., section handling).
        // Uncomment the following lines to use this approach instead.
        /*
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Ensure the format is set to Xlsx (required for XlsxSaveOptions).
            SaveFormat = SaveFormat.Xlsx,

            // Example: Save each document section to a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document with the specified options.
        doc.Save(outputFile, xlsxOptions);
        */
    }
}
