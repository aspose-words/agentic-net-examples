using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the shapes into a GroupShape and insert it at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optional: adjust group properties (size, wrap, etc.).
        group.WrapType = WrapType.None;
        group.Bounds = new RectangleF(0, 0, 300, 300);

        // Save the document to a file.
        string artifactsDir = "Artifacts/";
        System.IO.Directory.CreateDirectory(artifactsDir);
        doc.Save(artifactsDir + "GroupShape.docx");
    }
}
