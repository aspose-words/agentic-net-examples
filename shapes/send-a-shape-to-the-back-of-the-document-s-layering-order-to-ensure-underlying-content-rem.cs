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

        // Retrieve all shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Send the third shape to the back of the layering order.
        // Lower ZOrder values are rendered behind higher values.
        shapes[2].ZOrder = 0;

        // Optional validation: ensure the ZOrder was set.
        if (shapes[2].ZOrder != 0)
            throw new InvalidOperationException("Failed to set shape ZOrder.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeBackOrder.docx");
        doc.Save(outputPath);
    }
}
