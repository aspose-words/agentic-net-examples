using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        string docxPath = Path.Combine(outputDir, "sample.docx");
        string xlsxPath = Path.Combine(outputDir, "sample.xlsx");

        // -----------------------------------------------------------------
        // 1. Create a sample PDF containing a simple table.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        builder.Writeln("Sample PDF with a table:");

        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndTable();

        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Validate PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Validate DOCX creation.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the DOCX and convert it to XLSX (spreadsheet).
        // -----------------------------------------------------------------
        Document docxDoc = new Document(docxPath);

        // Use XlsxSaveOptions to control the conversion if needed.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            SaveFormat = SaveFormat.Xlsx,
            // Default behavior creates a separate worksheet per document section.
            // No additional configuration required for this example.
        };

        docxDoc.Save(xlsxPath, xlsxOptions);

        // Validate XLSX creation.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // All conversions completed successfully.
        Console.WriteLine("Conversion sequence completed:");
        Console.WriteLine($"PDF  -> {pdfPath}");
        Console.WriteLine($"DOCX -> {docxPath}");
        Console.WriteLine($"XLSX -> {xlsxPath}");
    }
}
