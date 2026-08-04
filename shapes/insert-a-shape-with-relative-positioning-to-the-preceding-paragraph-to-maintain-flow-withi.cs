using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeRelativePosition.docx");

        // Create a new empty document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph. The shape will be positioned relative to this paragraph.
        builder.Writeln("This is the first paragraph. It will be followed by a shape.");

        // Insert a floating rectangle shape.
        // The shape is positioned relative to the preceding paragraph (RelativeVerticalPosition.Paragraph)
        // and to the left margin (RelativeHorizontalPosition.Margin). No text wrapping (WrapType.None).
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin,   // Horizontal reference: left margin.
            0,                                   // Left offset (points) from the reference.
            RelativeVerticalPosition.Paragraph, // Vertical reference: the paragraph we just wrote.
            0,                                   // Top offset (points) from the reference.
            100,                                 // Width (points).
            50,                                  // Height (points).
            WrapType.None);                      // No wrapping; shape floats.

        // Apply a simple fill color to make the shape visible.
        shape.FillColor = System.Drawing.Color.LightBlue;

        // Write another paragraph after the shape to demonstrate flow continuation.
        builder.Writeln("This paragraph appears after the shape, maintaining the document flow.");

        // Save the document to the specified file.
        doc.Save(outputPath);

        // Validate that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // Optional: inform the user that the operation completed.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
