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

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape.
        group.Bounds = new RectangleF(0, 0, 200, 200);
        group.WrapType = WrapType.None; // Floating shape.
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.HorizontalAlignment = HorizontalAlignment.Center;
        group.VerticalAlignment = VerticalAlignment.Center;

        // Create a child shape (a rectangle) and add it to the group.
        Shape child = new Shape(doc, ShapeType.Rectangle);
        child.Width = 100;
        child.Height = 50;
        child.Left = 50; // Position inside the group.
        child.Top = 50;
        child.Fill.ForeColor = Color.LightBlue;
        child.Stroke.Color = Color.DarkBlue;
        group.AppendChild(child);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
