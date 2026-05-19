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

        // Insert three overlapping floating shapes.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100, 200, 200, WrapType.None);
        shape1.FillColor = System.Drawing.Color.Orange;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 150, 200, 200, WrapType.None);
        shape2.FillColor = System.Drawing.Color.LightBlue;

        Shape shape3 = builder.InsertShape(ShapeType.Triangle, RelativeHorizontalPosition.Page, 200,
            RelativeVerticalPosition.Page, 200, 200, 200, WrapType.None);
        shape3.FillColor = System.Drawing.Color.LightGreen;

        // Retrieve all shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

        // Send the first shape (orange rectangle) to the back of the layering order.
        // Lower ZOrder values are rendered behind higher values.
        shapes[0].ZOrder = 0;

        // Optional: ensure other shapes have higher ZOrder values.
        if (shapes.Length > 1) shapes[1].ZOrder = 2;
        if (shapes.Length > 2) shapes[2].ZOrder = 3;

        // Validate that the first shape is indeed at the back.
        if (shapes.Any(s => s != shapes[0] && s.ZOrder <= shapes[0].ZOrder))
            throw new InvalidOperationException("Failed to send the shape to the back of the layering order.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Shape.SendToBack.docx");
        doc.Save(outputPath);
    }
}
