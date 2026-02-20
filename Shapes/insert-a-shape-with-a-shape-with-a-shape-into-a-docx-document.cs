using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertNestedShapes
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the outermost group shape.
        GroupShape outerGroup = new GroupShape(doc);
        // Set the size and position of the outer group shape.
        outerGroup.Width = 400;
        outerGroup.Height = 300;
        outerGroup.Left = 50;
        outerGroup.Top = 50;
        // Add the outer group shape to the document.
        builder.CurrentParagraph.AppendChild(outerGroup);

        // Create a second (inner) group shape that will be placed inside the outer group.
        GroupShape innerGroup = new GroupShape(doc);
        // Set its size and position relative to the outer group.
        innerGroup.Width = 250;
        innerGroup.Height = 180;
        innerGroup.Left = 75;
        innerGroup.Top = 60;
        // Add the inner group shape to the outer group.
        outerGroup.AppendChild(innerGroup);

        // Create the innermost shape (a simple rectangle) and add it to the inner group.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        innerShape.Width = 150;
        innerShape.Height = 100;
        innerShape.Left = 50;
        innerShape.Top = 40;
        innerShape.FillColor = System.Drawing.Color.LightBlue;
        innerShape.StrokeColor = System.Drawing.Color.DarkBlue;
        innerShape.StrokeWeight = 2.0;
        // Add the innermost shape to the inner group.
        innerGroup.AppendChild(innerShape);

        // Save the document to a DOCX file.
        doc.Save("NestedShapes.docx");
    }
}
