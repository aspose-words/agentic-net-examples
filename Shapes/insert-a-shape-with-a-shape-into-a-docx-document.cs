using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape (in points).
        group.Width = 200;
        group.Height = 150;
        group.Left = 100;
        group.Top = 100;
        // Make the group floating so it can be positioned absolutely.
        group.WrapType = WrapType.None;

        // Create an inner rectangle shape.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        // Set size and position relative to the group shape.
        innerShape.Width = 180;
        innerShape.Height = 130;
        innerShape.Left = 10;   // Position inside the group.
        innerShape.Top = 10;
        // Optional: give the shape a fill color.
        innerShape.FillColor = System.Drawing.Color.LightBlue;
        // Add the inner shape to the group.
        group.AppendChild(innerShape);

        // Insert the group shape into the document.
        // The builder's current paragraph is the insertion point.
        builder.CurrentParagraph.AppendChild(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeWithInnerShape.docx");
    }
}
