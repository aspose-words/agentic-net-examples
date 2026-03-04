using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting Excel file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.xlsx";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Create Excel save options.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

        // Example customizations:
        // 1. Compress the XLSX file using maximum compression.
        xlsxOptions.CompressionLevel = CompressionLevel.Maximum;

        // 2. Save each Word section to a separate worksheet.
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // 3. Enable pretty formatting for better readability (optional).
        xlsxOptions.PrettyFormat = true;

        // Ensure the save format is set to Xlsx (required by the options object).
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an Excel file using the configured options.
        doc.Save(outputPath, xlsxOptions);
    }
}
