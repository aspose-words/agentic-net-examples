using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\SampleDocument.docx";

        // Path where the resulting Excel file will be saved.
        string outputPath = @"C:\Docs\SampleDocument.xlsx";

        // Load the Word document from the file system.
        Document doc = new Document(inputPath);

        // Configure save options for XLSX format.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Save each section of the Word document as a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets,

            // Explicitly set the format to XLSX (optional, but clarifies intent).
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as an Excel workbook using the configured options.
        doc.Save(outputPath, xlsxOptions);
    }
}
