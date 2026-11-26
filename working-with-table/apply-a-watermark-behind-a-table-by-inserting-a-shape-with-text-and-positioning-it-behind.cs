using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph to host the watermark shape.
        builder.Writeln();

        // Create a text shape that will act as a watermark.
        Shape watermark = new Shape(doc, ShapeType.TextPlainText);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        watermark.Width = 500;
        watermark.Height = 100;
        watermark.Rotation = -40;
        watermark.FillColor = System.Drawing.Color.LightGray;
        watermark.StrokeColor = System.Drawing.Color.LightGray;

        // Position the shape relative to the page and send it behind the text.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.WrapType = WrapType.None;
        watermark.BehindText = true;

        // Insert the watermark shape into the document.
        builder.InsertNode(watermark);

        // Build a simple table after the watermark.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "WatermarkTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
