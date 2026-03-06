using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file that contains tables.
        Document doc = new Document("InputDocument.docx");

        // Convert any table style formatting to direct formatting.
        // This ensures that the table appearance is preserved when exporting.
        doc.ExpandTableStylesToDirectFormatting();

        // Configure save options for XLSX output.
        // MultipleWorksheets creates a separate worksheet for each section of the document.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        saveOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Save the document as an XLSX workbook.
        doc.Save("OutputWorkbook.xlsx", saveOptions);
    }
}
