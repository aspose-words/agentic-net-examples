using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting XLSX file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.xlsx";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Configure save options for XLSX format.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Each section of the document will be saved as a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document as XLSX using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
