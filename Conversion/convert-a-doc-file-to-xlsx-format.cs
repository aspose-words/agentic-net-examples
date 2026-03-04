using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\Sample.doc";

        // Path where the XLSX file will be saved.
        string outputPath = @"C:\Docs\Sample.xlsx";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Option 1: Directly save using the SaveFormat enumeration.
        doc.Save(outputPath, SaveFormat.Xlsx);

        // Option 2: Save using XlsxSaveOptions for additional control (uncomment if needed).
        // XlsxSaveOptions options = new XlsxSaveOptions();
        // options.SectionMode = XlsxSectionMode.SingleWorksheet; // Example setting.
        // doc.Save(outputPath, options);
    }
}
