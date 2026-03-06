using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class ExportGroupShapeAsPng
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder for constructing content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two simple shapes.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        shape1.Stroke.Color = System.Drawing.Color.Red;
        shape1.FillColor = System.Drawing.Color.LightYellow;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        shape2.Stroke.Color = System.Drawing.Color.Blue;
        shape2.FillColor = System.Drawing.Color.LightGreen;

        // Group the two shapes into a GroupShape node.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // OPTIONAL: Adjust the group position if needed.
        group.Left = 100;
        group.Top = 100;

        // Locate the first GroupShape in the document.
        GroupShape targetGroup = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);
        if (targetGroup == null)
        {
            Console.WriteLine("No group shape found in the document.");
            return;
        }

        // Create a renderer for the group shape.
        ShapeRenderer renderer = targetGroup.GetShapeRenderer();

        // Define PNG save options (default format is PNG, but we set it explicitly).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Render the group shape to a PNG file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "GroupShape.png");
        renderer.Save(outputPath, pngOptions);

        Console.WriteLine($"Group shape exported to PNG: {outputPath}");
    }
}
