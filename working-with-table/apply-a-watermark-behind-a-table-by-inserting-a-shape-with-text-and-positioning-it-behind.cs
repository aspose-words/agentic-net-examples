using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a watermark shape behind the table content.
        InsertWatermark(builder, "CONFIDENTIAL");

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "WatermarkTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Output file was not created.");
    }

    private static void InsertWatermark(DocumentBuilder builder, string text)
    {
        // Position the builder at the start of the document.
        builder.MoveToDocumentStart();

        // Create a plain‑text shape that will serve as the watermark.
        Shape watermark = new Shape(builder.Document, ShapeType.TextPlainText);
        watermark.Width = 500;
        watermark.Height = 100;
        watermark.Rotation = -40; // Slight rotation for visual effect.
        watermark.WrapType = WrapType.None;
        watermark.BehindText = true; // Ensure the shape is behind other content.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.VerticalAlignment = VerticalAlignment.Center;
        watermark.HorizontalAlignment = HorizontalAlignment.Center;

        // Configure the watermark text.
        watermark.TextPath.Text = text;
        watermark.TextPath.FontFamily = "Arial";
        // FontSize property is not available in this version; rely on default size or adjust shape dimensions.
        watermark.Fill.Color = Color.LightGray;
        watermark.Stroke.Color = Color.LightGray;

        // Insert the shape into the document.
        builder.InsertNode(watermark);
    }
}
