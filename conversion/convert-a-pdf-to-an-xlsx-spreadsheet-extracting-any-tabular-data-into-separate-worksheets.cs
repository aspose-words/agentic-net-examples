using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class PdfToXlsxConverter
{
    public static void Main()
    {
        // Define file names
        const string pdfPath = "sample.pdf";
        const string xlsxPath = "output.xlsx";

        // Step 1: Create a sample document with a table and save it as PDF
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a simple table with two rows and two columns
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document as PDF (optional, just to have a PDF file)
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // Step 2: Convert the original document (which contains the table) to XLSX.
        // Aspose.Words extracts each table into a separate worksheet.
        sourceDoc.Save(xlsxPath, SaveFormat.Xlsx);

        // Validate that the XLSX file was created
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("Expected XLSX file was not created.");

        // Optional: Clean up generated files (comment out if you want to keep them)
        // File.Delete(pdfPath);
        // File.Delete(xlsxPath);
    }
}
