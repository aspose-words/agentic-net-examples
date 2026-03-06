using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Load the source document (any format supported by Aspose.Words)
        string inputPath = @"C:\Docs\SourceDocument.docx";
        Document doc = new Document(inputPath);

        // Create and configure Excel save options
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        // Save each section of the Word document as a separate worksheet
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;
        // Apply maximum compression to the resulting XLSX package
        xlsxOptions.CompressionLevel = CompressionLevel.Maximum;
        // Output pretty‑formatted XML inside the XLSX for readability
        xlsxOptions.PrettyFormat = true;
        // Explicitly set the format (required by the SaveOptions contract)
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an Excel workbook using the configured options
        string outputPath = @"C:\Docs\ResultWorkbook.xlsx";
        doc.Save(outputPath, xlsxOptions);
    }
}
