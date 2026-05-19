using System;
using System.IO;
using System.Linq;
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

        // Insert three overlapping rectangles. Newer shapes are placed on top by default.
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

        // Choose the first shape (orange rectangle) and read its current Z‑order.
        Shape targetShape = shapes[0];
        int originalZOrder = targetShape.ZOrder;
        Console.WriteLine($"Original ZOrder of target shape: {originalZOrder}");

        // Determine the highest ZOrder among the shapes.
        int maxZOrder = shapes.Max(s => s.ZOrder);

        // Bring the target shape to the front by assigning a higher ZOrder.
        targetShape.ZOrder = maxZOrder + 1;
        Console.WriteLine($"New ZOrder of target shape: {targetShape.ZOrder}");

        // Save the document to a local folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outFile = Path.Combine(artifactsDir, "ShapeZOrderDemo.docx");
        doc.Save(outFile);

        // Verify that the file was created.
        if (!File.Exists(outFile))
            throw new Exception("Document was not saved correctly.");

        Console.WriteLine($"Document saved to: {outFile}");
    }
}
