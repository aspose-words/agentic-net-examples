using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertShapeRelativeToParagraph
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text in the first paragraph.
        builder.Writeln("This is the first paragraph. The shape will be positioned relative to this paragraph.");

        // Insert a floating rectangle shape.
        // Position it relative to the preceding paragraph (vertical) and to the left margin (horizontal).
        // Left and Top are set to 0 so the shape starts exactly at the paragraph's anchor point.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,                     // Shape type.
            RelativeHorizontalPosition.Margin,       // Horizontal reference (left margin).
            0,                                       // Left offset in points.
            RelativeVerticalPosition.Paragraph,      // Vertical reference (the paragraph we just wrote).
            0,                                       // Top offset in points.
            100,                                     // Width in points.
            50,                                      // Height in points.
            WrapType.None);                          // No text wrapping; shape floats.

        // Optional visual styling.
        shape.StrokeColor = System.Drawing.Color.Blue;
        shape.FillColor = System.Drawing.Color.LightGray;

        // Add a second paragraph after the shape to demonstrate flow.
        builder.Writeln("This is the second paragraph, appearing after the floating shape.");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Inform that the process completed successfully (no interactive prompts).
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
