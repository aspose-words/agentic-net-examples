using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting XLSX file will be saved.
        string outputPath = "output.xlsx";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Create XLSX save options and explicitly set the format.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            // Optional: control how sections are mapped to worksheets.
            // SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document as an XLSX spreadsheet using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
