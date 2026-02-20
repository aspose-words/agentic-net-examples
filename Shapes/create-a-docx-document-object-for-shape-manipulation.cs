using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a rectangle shape.
        Shape shape = new Shape(doc, ShapeType.Rectangle);
        shape.Width = 100;               // Width in points.
        shape.Height = 50;               // Height in points.
        shape.Left = 100;                // Position from the left edge.
        shape.Top = 100;                 // Position from the top edge.
        shape.WrapType = WrapType.None;  // No text wrapping.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;
        shape.StrokeWeight = 2;          // Stroke thickness in points.

        // Insert the shape into the document.
        builder.InsertNode(shape);

        // Add a paragraph with text inside the shape.
        shape.AppendChild(new Paragraph(doc));
        shape.FirstParagraph.AppendChild(new Run(doc, "Hello Shape"));

        // Save the document with OOXML compliance that supports DML shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        doc.Save("ShapeManipulation.docx", saveOptions);
    }
}
