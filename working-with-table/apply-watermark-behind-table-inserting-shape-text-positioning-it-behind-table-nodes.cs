using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class WatermarkBehindTableExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table from the document body.
        table = doc.FirstSection.Body.Tables[0];

        // Create a floating shape that will serve as the text watermark.
        Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
        // Set the size of the shape (width, height) in points.
        watermarkShape.Width = 300;
        watermarkShape.Height = 100;

        // Configure the shape to behave as a floating object.
        watermarkShape.WrapType = WrapType.None;               // No text wrapping – shape floats.
        watermarkShape.BehindText = true;                      // Place shape behind the text.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermarkShape.HorizontalAlignment = HorizontalAlignment.Center;
        watermarkShape.VerticalAlignment = VerticalAlignment.Center;

        // Add the actual watermark text using the TextPath object.
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        // FontSize property is not available in older Aspose.Words versions; omit or set via shape scaling.
        watermarkShape.TextPath.Bold = true;
        watermarkShape.Rotation = -45;                         // Diagonal layout.

        // Optional visual styling.
        watermarkShape.Fill.Color = Color.LightGray;           // Light fill to make text visible.
        watermarkShape.StrokeColor = Color.Empty;              // No outline.

        // Insert the shape *before* the table so that it appears behind the table nodes.
        table.ParentNode.InsertBefore(watermarkShape, table);

        // Save the document.
        doc.Save("WatermarkBehindTable.docx");
    }
}
