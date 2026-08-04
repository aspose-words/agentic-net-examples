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

        // Insert three overlapping rectangles.
        // The newer shape is placed on top by default.
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 100,
            RelativeVerticalPosition.TopMargin, 100,
            200, 200,
            WrapType.None);
        shape1.FillColor = Color.Orange;

        Shape shape2 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 150,
            RelativeVerticalPosition.TopMargin, 150,
            200, 200,
            WrapType.None);
        shape2.FillColor = Color.LightBlue;

        Shape shape3 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 200,
            RelativeVerticalPosition.TopMargin, 200,
            200, 200,
            WrapType.None);
        shape3.FillColor = Color.LightGreen;

        // Retrieve all top‑level shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Display the current Z‑order values.
        Console.WriteLine("Current Z‑order values:");
        for (int i = 0; i < shapes.Length; i++)
        {
            Console.WriteLine($"Shape {i + 1}: ZOrder = {shapes[i].ZOrder}");
        }

        // Bring the first shape (orange rectangle) to the front.
        int maxZ = shapes.Max(s => s.ZOrder);
        shapes[0].ZOrder = maxZ + 1;

        // Verify the new Z‑order.
        Console.WriteLine("\nAfter bringing the first shape to front:");
        for (int i = 0; i < shapes.Length; i++)
        {
            Console.WriteLine($"Shape {i + 1}: ZOrder = {shapes[i].ZOrder}");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeZOrder.docx");
        doc.Save(outputPath);
        Console.WriteLine($"\nDocument saved to: {outputPath}");
    }
}
