using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsToExcelDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputFile = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting Excel file will be saved.
            string outputFile = @"C:\Docs\ConvertedWorkbook.xlsx";

            // Load the Word document from the file system.
            Document doc = new Document(inputFile);

            // Create XlsxSaveOptions to fine‑tune the conversion.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                // Save each Word section on a separate worksheet.
                SectionMode = XlsxSectionMode.MultipleWorksheets,

                // Use maximum compression for the resulting XLSX package.
                CompressionLevel = CompressionLevel.Maximum,

                // Ensure the format is explicitly set (optional, but clear).
                SaveFormat = SaveFormat.Xlsx
            };

            // Save the document as an Excel workbook using the configured options.
            doc.Save(outputFile, xlsxOptions);
        }
    }
}
