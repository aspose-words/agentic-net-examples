using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Input\Sample.doc";

        // Path where the XLSX file will be saved.
        string outputPath = @"C:\Output\Sample.xlsx";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Option 1: Directly save using the SaveFormat enumeration.
        doc.Save(outputPath, SaveFormat.Xlsx);

        // Option 2: Use XlsxSaveOptions for additional control (e.g., single worksheet mode).
        // Uncomment the following lines if you need custom save options.
        /*
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Ensure the format is set to Xlsx (required by the property).
            SaveFormat = SaveFormat.Xlsx,

            // Example: Save all sections to a single worksheet.
            SectionMode = XlsxSectionMode.SingleWorksheet
        };
        doc.Save(outputPath, xlsxOptions);
        */
    }
}
