using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOCX file that contains tables.
        string inputPath = @"C:\Input\DocumentWithTables.docx";

        // Path where the resulting XLSX workbook will be saved.
        string outputPath = @"C:\Output\TablesWorkbook.xlsx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Optional: expand any table style formatting to direct formatting.
        // This ensures that table appearance is preserved when converting.
        doc.ExpandTableStylesToDirectFormatting();

        // Configure XLSX save options.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Each section of the Word document will become a separate worksheet.
            // If the document has only one section, all tables will be placed on that worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets,

            // Explicitly set the save format to XLSX (optional, but clarifies intent).
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as an XLSX workbook using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
