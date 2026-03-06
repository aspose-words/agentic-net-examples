using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        GroupShapeToXps.Run();
    }
}

public class GroupShapeToXps
{
    public static void Run()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and stroke color.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and stroke color.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Prepare XPS save options (default settings are sufficient for this example).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the document, which now contains the group shape, as an XPS file.
        doc.Save("GroupShape.xps", xpsOptions);
    }
}
