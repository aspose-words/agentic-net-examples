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

        // First paragraph – the shape will be anchored to this paragraph.
        builder.Writeln("This is the first paragraph. The shape will appear after it.");

        // Insert a floating rectangle shape.
        // Position: relative to the preceding paragraph (vertical) and to the left margin (horizontal).
        // Left = 0, Top = 0 means the shape starts at the paragraph's left edge.
        // Width = 100 points, Height = 50 points.
        // WrapType.TopBottom makes the text flow above and below the shape.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin, 0,
            RelativeVerticalPosition.Paragraph, 0,
            100, 50,
            WrapType.TopBottom);

        // Optional visual styling.
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;
        shape.StrokeWeight = 1.5;

        // Continue with another paragraph to demonstrate text flow around the shape.
        builder.Writeln("This is the second paragraph. It should appear after the floating shape, respecting the wrap.");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputShape.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
