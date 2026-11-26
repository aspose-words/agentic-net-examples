using System;
using System.Drawing;
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

        // Add some introductory text.
        builder.Writeln("This paragraph appears before the shape. It contains enough text to demonstrate wrapping around the shape on both sides of the page. The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog.");

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Configure the shape to wrap text on both sides of the page.
        shape.WrapType = WrapType.Square;               // Enable wrapping.
        shape.WrapSide = WrapSide.Both;                 // Wrap on both left and right sides.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = ConvertUtil.MillimeterToPoint(20); // Position from the left edge of the page.
        shape.Top = ConvertUtil.MillimeterToPoint(20);  // Position from the top edge of the page.
        shape.StrokeColor = Color.Blue;                // Outline color.
        shape.FillColor = Color.LightGray;              // Fill color.

        // Add more text after the shape to show wrapping.
        builder.Writeln();
        builder.Writeln("This paragraph appears after the shape. It also contains sufficient text to wrap around the shape on both sides. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to disk.
        string outputPath = "WrapShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
