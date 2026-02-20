using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new group shape.
        GroupShape group = new GroupShape(doc);
        // Set the size of the group.
        group.Width = 300;
        group.Height = 200;
        // Position the group on the page.
        group.WrapType = WrapType.None;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.Left = 100;
        group.Top = 100;

        // Create a rectangle shape to be a child of the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 150;
        rect.Height = 100;
        rect.Left = 0;
        rect.Top = 0;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.Blue;

        // Create an ellipse shape to be a child of the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 100;
        ellipse.Height = 100;
        ellipse.Left = 150;
        ellipse.Top = 50;
        ellipse.Fill.ForeColor = Color.LightCoral;
        ellipse.Stroke.Color = Color.Red;

        // Add the child shapes to the group.
        group.AppendChild(rect);
        group.AppendChild(ellipse);

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the modified document back to HTML (or any other format you need).
        doc.Save("output.html");
    }
}
