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

        // Add a paragraph with some text.
        builder.Writeln("This is the first paragraph.");

        // Insert a rectangle shape after the paragraph.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Configure the shape to be positioned relative to the preceding paragraph.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;   // Relative to page margin.
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;   // Relative to the paragraph.
        shape.WrapType = WrapType.Square;                                      // Text wraps around the shape.
        shape.Left = 0;    // Align with the left margin.
        shape.Top = 0;     // Align with the top of the paragraph.

        // Optional visual styling.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Save the document.
        string outputPath = "ShapeRelativePosition.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends here without waiting for user input.
    }
}
