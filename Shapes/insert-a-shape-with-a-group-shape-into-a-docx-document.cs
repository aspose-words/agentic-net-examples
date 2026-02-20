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
        // Define the size and position of the group shape on the page.
        group.Width = 200;   // width in points
        group.Height = 100;  // height in points
        group.Left = 100;    // distance from the left edge of the page
        group.Top = 100;     // distance from the top edge of the page
        group.WrapType = WrapType.None; // make the group floating

        // Create a rectangle shape to be placed inside the group.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 80;
        rectangle.Height = 60;
        rectangle.Left = 0;   // position relative to the group's coordinate origin
        rectangle.Top = 0;
        rectangle.Fill.ForeColor = Color.LightBlue;
        rectangle.Stroke.Color = Color.DarkBlue;

        // Add the rectangle to the group shape.
        group.AppendChild(rectangle);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
