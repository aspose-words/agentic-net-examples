using System;
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

        // Insert a shape that will act as a watermark.
        Shape watermark = new Shape(doc, ShapeType.TextPlainText);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        watermark.Width = 500;
        watermark.Height = 100;
        watermark.Rotation = -40;
        watermark.WrapType = WrapType.None;          // No text wrapping.
        watermark.BehindText = true;                 // Place behind all text (including tables).
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.Left = 0;
        watermark.Top = 0;

        // Insert the watermark shape into the document.
        builder.InsertNode(watermark);

        // Add a paragraph after the watermark to host the table.
        builder.Writeln();

        // Build a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "WatermarkTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends here.
    }
}
