using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToExcelExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting Excel file will be saved.
            string outputPath = @"C:\Docs\ConvertedDocument.xlsx";

            // Load the Word document from disk.
            Document doc = new Document(inputPath);

            // Create save options for XLSX format.
            XlsxSaveOptions saveOptions = new XlsxSaveOptions();

            // Optional: Save each section of the Word document to a separate worksheet.
            // Uncomment the following line to enable this behavior.
            // saveOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            // Save the document as an Excel file using the specified options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
