using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and appearance.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        shape1.Left = 50;
        shape1.Top = 50;
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and appearance.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        shape2.Left = 300;
        shape2.Top = 80;
        shape2.Stroke.Color = Color.Blue;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Prepare XPS save options (default options are sufficient for this example).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the document, which now contains the group shape, as an XPS file.
        doc.Save("GroupShapeDocument.xps", xpsOptions);
    }
}
