using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape with specific dimensions.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        shape.StrokeColor = Color.Blue;
        shape.FillColor = Color.LightGray;

        // Ensure layout is up‑to‑date before retrieving actual bounds.
        doc.UpdatePageLayout();

        // Retrieve the actual bounds of the shape.
        RectangleF bounds = shape.Bounds;

        // Log the coordinate points.
        Console.WriteLine($"Actual bounds: X={bounds.X}, Y={bounds.Y}, Width={bounds.Width}, Height={bounds.Height}");

        // Save the document.
        string outputPath = "ShapeBounds.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
