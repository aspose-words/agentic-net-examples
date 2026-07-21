using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three overlapping rectangles.
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 100,
            RelativeVerticalPosition.TopMargin, 100,
            200, 200,
            WrapType.None);
        shape1.FillColor = System.Drawing.Color.Orange;

        Shape shape2 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 150,
            RelativeVerticalPosition.TopMargin, 150,
            200, 200,
            WrapType.None);
        shape2.FillColor = System.Drawing.Color.LightBlue;

        Shape shape3 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 200,
            RelativeVerticalPosition.TopMargin, 200,
            200, 200,
            WrapType.None);
        shape3.FillColor = System.Drawing.Color.LightGreen;

        // Retrieve all top‑level shapes.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Output the original ZOrder of the first shape.
        Console.WriteLine($"Original ZOrder of first shape: {shapes[0].ZOrder}");

        // Bring the first shape to the front by assigning it the highest ZOrder.
        int maxZOrder = shapes.Max(s => s.ZOrder);
        shapes[0].ZOrder = maxZOrder + 1;

        // Verify the change.
        Console.WriteLine($"New ZOrder of first shape: {shapes[0].ZOrder}");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeZOrderExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the document.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
