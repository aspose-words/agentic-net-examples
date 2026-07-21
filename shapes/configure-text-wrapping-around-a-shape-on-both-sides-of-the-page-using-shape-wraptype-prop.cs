using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ConfigureShapeWrapping
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text.
        builder.Writeln("This paragraph appears before the shape. It will demonstrate how text wraps around the shape on both sides of the page.");

        // Insert a floating rectangle shape with square text wrapping.
        // The shape is positioned 100 points from the left and top edges of the page.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            150, 150,
            WrapType.Square);

        // Ensure the shape uses the default WrapSide (Both) so text wraps on both sides.
        shape.WrapSide = WrapSide.Both;

        // Add more text after the shape to see the wrapping effect.
        builder.Writeln("This paragraph follows the shape. The text should flow around the shape on the left and right sides, demonstrating the Square wrap type with both-side wrapping.");

        // Define the output file path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeWrapBothSides.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
