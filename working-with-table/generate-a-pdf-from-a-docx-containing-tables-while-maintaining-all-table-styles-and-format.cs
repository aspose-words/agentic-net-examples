using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string docxPath = Path.Combine(outputDir, "SampleWithTable.docx");
        string pdfPath = Path.Combine(outputDir, "SampleWithTable.pdf");

        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header row and two data rows
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row
        builder.InsertCell();
        builder.Font.Bold = false;
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table
        builder.EndTable();

        // Create a custom table style and apply it to the table
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.0;
        customStyle.RowStripe = 2; // Enable row banding
        table.Style = customStyle;

        // Convert style formatting to direct formatting before PDF conversion
        doc.ExpandTableStylesToDirectFormatting();

        // Save the document as DOCX (sample source)
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX and convert it to PDF while preserving table formatting
        Document pdfDoc = new Document(docxPath);
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed: output file not found.");
    }
}
