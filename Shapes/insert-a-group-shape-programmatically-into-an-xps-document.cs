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

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's position or size here.
        // For example, set the group to be centered on the page.
        // group.Left = 100;
        // group.Top = 100;

        // Save the document as an XPS file.
        doc.Save("GroupShape.xps", SaveFormat.Xps);
    }
}
