using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three overlapping rectangles.
        // The newer shape is placed on top by default.
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

        // Retrieve all shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .ToArray();

        // Send the green rectangle to the back of the stacking order.
        // Lower ZOrder values are rendered behind higher values.
        // Here we set the green shape to the lowest value (0).
        shapes[2].ZOrder = 0;

        // Optionally adjust the other shapes' ZOrder to ensure they appear above the green one.
        shapes[0].ZOrder = 2; // Orange on top.
        shapes[1].ZOrder = 1; // Blue in the middle.

        // Save the resulting document.
        doc.Save("ShapeSendToBack.docx");
    }
}
