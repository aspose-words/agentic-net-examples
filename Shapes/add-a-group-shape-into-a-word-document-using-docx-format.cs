using System;
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

        // Insert two floating shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        shape1.Left = 50;   // Position relative to the page.
        shape1.Top = 50;
        shape1.Stroke.Color = System.Drawing.Color.Blue;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        shape2.Left = 300;
        shape2.Top = 100;
        shape2.Stroke.Color = System.Drawing.Color.Green;

        // Group the two shapes into a new GroupShape node and insert it at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally adjust the group's properties (e.g., set a title).
        group.Title = "My Group Shape";

        // Save the document in DOCX format.
        doc.Save("GroupShape.docx", SaveFormat.Docx);
    }
}
