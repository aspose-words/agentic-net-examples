using System;
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

        // Insert a rectangle shape positioned at the top of the page.
        Shape topShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        ConfigureFloatingShape(topShape, 0, 0);

        // Insert a rectangle shape positioned in the middle (approximate) of the page.
        Shape middleShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        ConfigureFloatingShape(middleShape, 0, 200); // Adjust the top offset as needed.

        // Insert a rectangle shape positioned at the bottom of the page.
        Shape bottomShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        ConfigureFloatingShape(bottomShape, 0, 400); // Adjust the top offset as needed.

        // Save the document.
        string outputPath = "VerticalAnchorShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
    }

    // Helper method to configure a shape as a floating shape with the specified position.
    private static void ConfigureFloatingShape(Shape shape, double left, double top)
    {
        // Make the shape floating.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Set absolute position on the page.
        shape.Left = left;
        shape.Top = top;
    }
}
