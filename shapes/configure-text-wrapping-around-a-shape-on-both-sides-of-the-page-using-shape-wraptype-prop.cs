using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed non risus. Suspendisse lectus tortor, dignissim sit amet, " +
                        "adipiscing nec, ultricies sed, dolor.");

        // Insert a floating rectangle shape with Square wrap (text on both sides).
        double left = 100;   // distance from the left margin (points)
        double top = 50;     // distance from the top margin (points)
        double width = 150;  // shape width (points)
        double height = 100; // shape height (points)

        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin, left,
            RelativeVerticalPosition.Margin, top,
            width, height,
            WrapType.Square); // WrapType.Square wraps text around all sides.

        // Optional visual styling.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Add more text that will wrap around the shape.
        builder.Writeln("Continuing the paragraph after the shape. " +
                        "The quick brown fox jumps over the lazy dog. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Praesent commodo cursus magna, vel scelerisque nisl consectetur et.");

        // Save the document.
        string outputPath = "ShapeWrapBothSides.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"The output file was not created: {outputPath}");
    }
}
