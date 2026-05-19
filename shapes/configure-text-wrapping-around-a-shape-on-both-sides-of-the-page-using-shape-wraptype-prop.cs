using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WrapShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the shape.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Praesent vitae eros eget tellus tristique bibendum. " +
                        "Donec rutrum sed sem quis venenatis.");

        // Insert a floating rectangle shape with Square wrapping (text wraps on both sides).
        // Position the shape 100 points from the left and 100 points from the top of the page.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            150, // width
            100, // height
            WrapType.Square);

        // Optional: give the shape a visible fill and outline.
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;
        shape.StrokeWeight = 1.5;

        // Write more text after the shape to demonstrate wrapping on the right side.
        builder.Writeln();
        builder.Writeln("Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                        "nisi ut aliquip ex ea commodo consequat.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WrapShape.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved successfully.");

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
