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
        Shape watermarkShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 300,
            Height = 100,
            WrapType = WrapType.None,               // No text wrapping – shape floats.
            BehindText = true,                      // Place shape behind the text.
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Rotation = -45                           // Diagonal layout.
        };

        // Add the actual watermark text using the TextPath object.
        watermarkShape.TextPath.Text = "CONFIDENTIAL";
        watermarkShape.TextPath.FontFamily = "Arial";
        watermarkShape.TextPath.Bold = true;

        // Optional visual styling.
        watermarkShape.Fill.Color = Color.LightGray;   // Light fill to make text visible.
        watermarkShape.StrokeColor = Color.Empty;      // No outline.

        // Insert the shape into a paragraph and place that paragraph before the table.
        Paragraph watermarkParagraph = new Paragraph(doc);
        watermarkParagraph.AppendChild(watermarkShape);
        table.ParentNode.InsertBefore(watermarkParagraph, table);

        // Save the document.
        doc.Save("WatermarkBehindTable.docx");
    }
}
