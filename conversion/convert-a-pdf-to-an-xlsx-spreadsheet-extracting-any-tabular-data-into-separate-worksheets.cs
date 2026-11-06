using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with a table.
        Document wordDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(wordDoc);

        // Build a simple table with two columns and two data rows.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the Word document as a PDF file.
        const string pdfPath = "sample.pdf";
        wordDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF using Aspose.Words.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to an XLSX spreadsheet.
        const string xlsxPath = "output.xlsx";
        pdfDoc.Save(xlsxPath, SaveFormat.Xlsx);

        // Verify the XLSX file was created.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // Optional cleanup (commented out).
        // File.Delete(pdfPath);
    }
}
