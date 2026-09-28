using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeLayeringExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text that will be overlapped by the shape.
        builder.Writeln("This is some sample text that will be overlapped by a shape.");

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Configure the shape to be floating so it can overlap the text.
        shape.WrapType = WrapType.None;                     // No text wrapping.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        shape.Left = 0;                                     // Position at the left margin.
        shape.Top = 0;                                      // Position at the top margin.

        // Send the shape to the back of the layering order.
        shape.BehindText = true; // Places the shape behind the text and other shapes.

        // Save the document.
        string outputPath = "ShapeBack.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
