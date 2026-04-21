using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three overlapping floating rectangles.
        // First rectangle (orange).
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 100,
            RelativeVerticalPosition.TopMargin, 100,
            200, 200,
            WrapType.None);
        shape1.FillColor = Color.Orange;

        // Second rectangle (light blue) overlapping the first.
        Shape shape2 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 150,
            RelativeVerticalPosition.TopMargin, 150,
            200, 200,
            WrapType.None);
        shape2.FillColor = Color.LightBlue;

        // Third rectangle (light green) overlapping the second.
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

        // Output current ZOrder values (for demonstration purposes).
        Console.WriteLine("Current ZOrder values:");
        for (int i = 0; i < shapes.Length; i++)
        {
            Console.WriteLine($"Shape {i}: ZOrder = {shapes[i].ZOrder}");
        }

        // Bring the first shape (orange rectangle) to the front.
        // Determine the highest existing ZOrder and set the target shape higher than that.
        int maxZOrder = shapes.Max(s => s.ZOrder);
        shapes[0].ZOrder = maxZOrder + 1;

        // Verify that the ZOrder was updated.
        if (shapes[0].ZOrder != maxZOrder + 1)
            throw new InvalidOperationException("Failed to update ZOrder.");

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeZOrder.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
