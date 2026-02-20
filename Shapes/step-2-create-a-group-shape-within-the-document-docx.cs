using Aspose.Words;
using Aspose.Words.Drawing;
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will hold other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size of the group shape (points).
        group.Width = 200;
        group.Height = 100;
        // Position the group shape on the page.
        group.Left = 50;
        group.Top = 50;
        // Define how text wraps around the group shape (None = floating).
        group.WrapType = WrapType.None;

        // Insert the group shape into the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Add a rectangle shape inside the group as an example child shape.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        innerShape.Width = 80;
        innerShape.Height = 40;
        innerShape.Left = 10;
        innerShape.Top = 10;
        group.AppendChild(innerShape);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
