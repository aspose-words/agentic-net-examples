using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text shape that will act as a watermark.
        Shape watermark = new Shape(doc, ShapeType.TextPlainText);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        // FontSize property is not available in this version; size is controlled by shape dimensions.
        watermark.Width = 500;
        watermark.Height = 100;
        watermark.Rotation = -40;
        watermark.FillColor = Color.LightGray;
        watermark.StrokeColor = Color.LightGray;
        watermark.WrapType = WrapType.None;          // No text wrapping.
        watermark.BehindText = true;                 // Place behind the table.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.Left = 0;
        watermark.Top = 0;

        // Insert the watermark shape into the document.
        builder.InsertNode(watermark);

        // Build a simple table after the watermark.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "WatermarkTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Confirmation message.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
