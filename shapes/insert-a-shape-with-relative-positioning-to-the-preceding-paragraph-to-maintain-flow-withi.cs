using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will serve as the anchor for the floating shape.
        builder.Writeln("This is the first paragraph. The shape will be positioned relative to this paragraph.");

        // Insert a floating rectangle shape anchored to the preceding paragraph.
        // The shape is positioned relative to the paragraph (vertical) and the left margin (horizontal).
        // Top and left offsets are set to 0 to place the shape directly after the paragraph.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin,   // Horizontal reference.
            0,                                   // Left offset.
            RelativeVerticalPosition.Paragraph, // Vertical reference (the paragraph we just wrote).
            0,                                   // Top offset.
            100,                                 // Width in points.
            50,                                  // Height in points.
            WrapType.None);                      // No text wrapping; shape stays in the flow.

        // Optional formatting to make the shape visible.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.Black;

        // Continue with more text to demonstrate that the shape stays in flow.
        builder.Writeln("This is the second paragraph, appearing after the shape.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RelativeShapeExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document.");

        // Optionally, inform that the process completed (no console input required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
