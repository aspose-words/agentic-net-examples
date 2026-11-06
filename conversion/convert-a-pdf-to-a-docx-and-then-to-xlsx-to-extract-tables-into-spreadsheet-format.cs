using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with a table and save it as PDF.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Insert a simple table.
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

        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the sample document as PDF.
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // Step 2: Load the PDF and convert it to DOCX.
        Document pdfDoc = new Document(pdfPath);
        string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Expected DOCX file was not created.");

        // Step 3: Load the DOCX and convert it to XLSX (Excel) to extract tables.
        Document docxDoc = new Document(docxPath);
        string xlsxPath = "tables.xlsx";
        docxDoc.Save(xlsxPath, SaveFormat.Xlsx);

        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("Expected XLSX file was not created.");

        // Program completed successfully.
    }
}
