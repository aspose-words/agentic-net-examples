using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words).
            string inputFile = @"C:\Docs\SourceDocument.docx";

            // Path where the Excel file will be saved.
            string outputFile = @"C:\Docs\ConvertedDocument.xlsx";

            // Load the source document using the Document constructor (lifecycle rule).
            Document doc = new Document(inputFile);

            // Optionally configure Excel-specific save options.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
            {
                // Ensure the format is set to Xlsx (required by the options object).
                SaveFormat = SaveFormat.Xlsx,

                // Example: save each section of the Word document to a separate worksheet.
                SectionMode = XlsxSectionMode.MultipleWorksheets
            };

            // Save the document as an Excel file using the Save method (lifecycle rule).
            doc.Save(outputFile, xlsxOptions);
        }
    }
}
