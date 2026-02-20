using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToExcelSample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the resulting Excel file will be saved.
            string outputPath = @"C:\Docs\SampleDocument.xlsx";

            // Load the Word document.
            Document doc = new Document(inputPath);

            // Configure save options for XLSX format.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                // Save each section of the Word document to a separate worksheet.
                SectionMode = XlsxSectionMode.MultipleWorksheets,

                // Optional: keep the default settings for other options.
                // For example, you could enable pretty formatting:
                // PrettyFormat = true
            };

            // Save the document as an Excel workbook using the configured options.
            doc.Save(outputPath, xlsxOptions);
        }
    }
}
