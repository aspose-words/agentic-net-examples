using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Input document path – can be any format supported by Aspose.Words (e.g., .docx, .pdf, .rtf, etc.).
            string inputPath = @"C:\Input\SampleDocument.docx";

            // Output Excel file path – the extension determines the Excel format.
            // Here we use .xlsx (Office Open XML Spreadsheet) which is fully supported.
            string outputPath = @"C:\Output\ConvertedDocument.xlsx";

            // Load the source document using the Document constructor that accepts a file name.
            Document doc = new Document(inputPath);

            // Configure XLSX save options.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                // Explicitly set the save format to Xlsx (required by the options class).
                SaveFormat = SaveFormat.Xlsx,

                // Example: save each section of the source document to a separate worksheet.
                SectionMode = XlsxSectionMode.MultipleWorksheets
            };

            // Save the document to the Excel format using the Save method that accepts a file name and SaveOptions.
            doc.Save(outputPath, xlsxOptions);

            Console.WriteLine("Document successfully converted to Excel format.");
        }
    }
}
