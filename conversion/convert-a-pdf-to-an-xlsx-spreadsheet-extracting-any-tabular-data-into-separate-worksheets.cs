using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the sample PDF and the resulting XLSX.
        string pdfPath = "sample.pdf";
        string xlsxPath = "output.xlsx";

        // Create a sample PDF containing two tables.
        CreateSamplePdf(pdfPath);

        // Load the PDF with Aspose.Words.
        Document pdfDocument = new Document(pdfPath);

        // Save the document as XLSX. Each table will be placed on a separate worksheet.
        pdfDocument.Save(xlsxPath, SaveFormat.Xlsx);

        // Validate that the XLSX file was created and is not empty.
        if (!File.Exists(xlsxPath) || new FileInfo(xlsxPath).Length == 0)
        {
            throw new InvalidOperationException("Expected output XLSX was not created or is empty.");
        }

        Console.WriteLine($"Conversion succeeded. XLSX saved to '{xlsxPath}'.");
    }

    private static void CreateSamplePdf(string path)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First table (3 columns, 4 rows including header).
        builder.StartTable();
        // Header row.
        builder.InsertCell();
        builder.Write("Header1");
        builder.InsertCell();
        builder.Write("Header2");
        builder.InsertCell();
        builder.Write("Header3");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"R{i}C1");
            builder.InsertCell();
            builder.Write($"R{i}C2");
            builder.InsertCell();
            builder.Write($"R{i}C3");
            builder.EndRow();
        }
        builder.EndTable();

        // Insert a page break before the second table.
        builder.InsertBreak(BreakType.PageBreak);

        // Second table (2 columns, 3 rows including header).
        builder.StartTable();
        // Header row.
        builder.InsertCell();
        builder.Write("ColA");
        builder.InsertCell();
        builder.Write("ColB");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Write($"A{i}");
            builder.InsertCell();
            builder.Write($"B{i}");
            builder.EndRow();
        }
        builder.EndTable();

        // Save the document as PDF.
        doc.Save(path, SaveFormat.Pdf);
    }
}
