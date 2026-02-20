using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoHtmlFixed
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will hold other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size of the group shape (in points).
        group.Width = 200;
        group.Height = 100;
        // Position the group shape on the page.
        group.Left = 100;
        group.Top = 100;
        // Make the group shape floating (not inline).
        group.WrapType = WrapType.None;

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 180;
        rect.Height = 80;
        rect.Left = 10;   // Position relative to the group's top‑left corner.
        rect.Top = 10;
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 1.5;

        // Add the rectangle to the group shape.
        group.AppendChild(rect);

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the document in HTML Fixed format.
        doc.Save("GroupShape.html", SaveFormat.HtmlFixed);
    }
}
