using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape that will hold other shapes.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        // Here we set a 200x200 points rectangle at the top‑left corner of the page.
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Set the group shape to be a floating shape (no text wrapping).
        group.WrapType = WrapType.None;
        group.BehindText = true;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Create a child shape (a simple rectangle) to place inside the group.
        Shape childShape = new Shape(doc, ShapeType.Rectangle);
        childShape.Width = 100;   // width in points
        childShape.Height = 50;   // height in points
        childShape.Left = 20;     // position relative to the group's coordinate space
        childShape.Top = 20;
        childShape.WrapType = WrapType.None;
        childShape.Fill.Color = Color.LightBlue;

        // Add the child shape to the group.
        group.AppendChild(childShape);

        // Save the document as DOCX.
        doc.Save("GroupShapeLayout.docx");
    }
}
