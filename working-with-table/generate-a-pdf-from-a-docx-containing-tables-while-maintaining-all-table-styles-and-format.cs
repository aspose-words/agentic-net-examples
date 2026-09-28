using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with a styled table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row - header cells.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Bold = true;
        builder.Write("Header 1");

        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Bold = true;
        builder.Write("Header 2");

        builder.EndRow();

        // Second row - data cells.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Font.Bold = false;
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Write("Cell 2");

        builder.EndRow();

        builder.EndTable();

        // Apply a built‑in table style to preserve formatting.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.StyleIdentifier = StyleIdentifier.LightListAccent1;
        table.AllowAutoFit = true;

        // Save the sample DOCX locally.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleTable.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX and convert it to PDF, preserving table styles.
        Document loadedDoc = new Document(docxPath);
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleTable.pdf");
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new Exception("PDF conversion failed: output file not found.");
        }
    }
}
