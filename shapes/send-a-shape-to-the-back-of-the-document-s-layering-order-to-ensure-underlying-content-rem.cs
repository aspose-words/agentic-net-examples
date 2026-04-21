using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class ShapeLayeringExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three overlapping rectangles. The later inserted shape is on top by default.
        Shape orangeRect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 100,
            RelativeVerticalPosition.TopMargin, 100,
            200, 200,
            WrapType.None);
        orangeRect.FillColor = Color.Orange;

        Shape blueRect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 150,
            RelativeVerticalPosition.TopMargin, 150,
            200, 200,
            WrapType.None);
        blueRect.FillColor = Color.LightBlue;

        Shape greenRect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 200,
            RelativeVerticalPosition.TopMargin, 200,
            200, 200,
            WrapType.None);
        greenRect.FillColor = Color.LightGreen;

        // Retrieve all shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Set ZOrder values so that the orange rectangle is at the back.
        // Lower ZOrder means farther back; higher means in front.
        shapes[0].ZOrder = 1; // orange
        shapes[1].ZOrder = 2; // blue
        shapes[2].ZOrder = 3; // green

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeLayering.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // No interactive prompts; program ends here.
    }
}
