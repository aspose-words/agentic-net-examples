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
            string inputPath = @"C:\Docs\ExampleDocument.docx";

            // Path where the resulting Excel file will be saved.
            string outputPath = @"C:\Docs\ExampleDocument.xlsx";

            // Load the Word document.
            Document doc = new Document(inputPath);

            // Create XlsxSaveOptions to control how the document is saved as XLSX.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

            // Optional: Save each section of the Word document to a separate worksheet.
            xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            // Ensure the SaveFormat is set to Xlsx (required by the options object).
            xlsxOptions.SaveFormat = SaveFormat.Xlsx;

            // Save the document as an Excel workbook using the specified options.
            doc.Save(outputPath, xlsxOptions);
        }
    }
}
