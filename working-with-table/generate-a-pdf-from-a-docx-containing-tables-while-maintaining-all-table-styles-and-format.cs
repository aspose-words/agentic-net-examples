using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "SampleTable.docx");
        string pdfPath = Path.Combine(outputDir, "SampleTable.pdf");

        // Create a new blank document and a builder to populate it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style and apply it to the table.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;
        customStyle.Borders.Color = Color.Blue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        table.Style = customStyle;

        // Convert any table style formatting to direct formatting.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the document as DOCX (optional source file).
        doc.Save(docxPath, SaveFormat.Docx);

        // Save the same document as PDF, preserving all table formatting.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the output files were created.
        if (!File.Exists(docxPath) || !File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Failed to create output files.");
        }
    }
}
