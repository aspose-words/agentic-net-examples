using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentToExcelConverter
{
    // Converts any supported document format to an Excel workbook (XLSX).
    public static void ConvertToExcel(string inputFilePath, string outputFilePath)
    {
        // Load the source document. The constructor automatically detects the file format.
        Document doc = new Document(inputFilePath);

        // Set up XLSX save options.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Explicitly specify the XLSX format.
            SaveFormat = SaveFormat.Xlsx,
            // Example: each section of the source document becomes a separate worksheet.
            SectionMode = XlsxSectionMode.MultipleWorksheets
        };

        // Save the document as an Excel file using the configured options.
        doc.Save(outputFilePath, saveOptions);
    }

    // Sample entry point demonstrating usage.
    public static void Main()
    {
        string inputPath = @"C:\Docs\Sample.docx";   // Can be DOCX, PDF, HTML, etc.
        string outputPath = @"C:\Docs\Sample.xlsx";

        ConvertToExcel(inputPath, outputPath);

        Console.WriteLine("Document successfully converted to Excel.");
    }
}
