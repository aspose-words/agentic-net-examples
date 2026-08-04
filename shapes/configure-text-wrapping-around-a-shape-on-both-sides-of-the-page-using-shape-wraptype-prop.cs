using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Define shape size.
        double shapeWidth = 100;
        double shapeHeight = 100;

        // Position the shape roughly in the center of the page.
        double left = (builder.PageSetup.PageWidth - shapeWidth) / 2;
        double top = (builder.PageSetup.PageHeight - shapeHeight) / 2;

        // Insert a floating rectangle with Square wrap type (wraps on both sides).
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, left,
            RelativeVerticalPosition.Page, top,
            shapeWidth, shapeHeight,
            WrapType.Square);

        // Make the shape visible.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Add more text that will wrap around the shape.
        builder.Writeln("\nMore text that should wrap around the shape on both sides. " +
                        "The quick brown fox jumps over the lazy dog. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeWrapBothSides.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
