using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three overlapping rectangles. The later inserted shape is on top by default.
        Shape orange = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 100,
            RelativeVerticalPosition.TopMargin, 100,
            200, 200, WrapType.None);
        orange.FillColor = Color.Orange;

        Shape blue = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 150,
            RelativeVerticalPosition.TopMargin, 150,
            200, 200, WrapType.None);
        blue.FillColor = Color.LightBlue;

        Shape green = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 200,
            RelativeVerticalPosition.TopMargin, 200,
            200, 200, WrapType.None);
        green.FillColor = Color.LightGreen;

        // Retrieve all top‑level shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Display original ZOrder values (for debugging purposes).
        Console.WriteLine("Original ZOrder values:");
        foreach (Shape s in shapes)
            Console.WriteLine($"{s.ShapeType}: {s.ZOrder}");

        // Bring the green rectangle to the front by assigning it the highest ZOrder.
        int maxZ = shapes.Max(s => s.ZOrder);
        green.ZOrder = maxZ + 1; // Higher value means frontmost.

        // Verify the new ZOrder values.
        Console.WriteLine("\nAfter bringing green shape to front:");
        foreach (Shape s in shapes)
            Console.WriteLine($"{s.ShapeType}: {s.ZOrder}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ZOrderDemo.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
