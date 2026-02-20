using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path to the destination XLSX file.
        string outputPath = @"C:\Docs\output.xlsx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Optionally configure how sections are saved.
        // Here we keep the default behavior (each section becomes a separate worksheet).
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Uncomment the line below to force all sections onto a single worksheet.
            // SectionMode = XlsxSectionMode.SingleWorksheet
        };

        // Save the document as an XLSX spreadsheet.
        doc.Save(outputPath, saveOptions);
    }
}
