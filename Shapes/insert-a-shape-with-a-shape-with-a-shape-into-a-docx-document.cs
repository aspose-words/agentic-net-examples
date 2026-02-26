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

        // Insert an outer group shape. No shapes are passed, so the group is empty initially.
        GroupShape outerGroup = builder.InsertGroupShape();

        // Create an inner group shape manually (floating by default).
        GroupShape innerGroup = new GroupShape(doc);
        innerGroup.Width = 150;   // Width in points.
        innerGroup.Height = 150;  // Height in points.
        innerGroup.Left = 20;     // Position relative to the outer group.
        innerGroup.Top = 20;
        innerGroup.WrapType = WrapType.Inline; // Make it inline for simplicity.

        // Add the inner group shape to the outer group.
        outerGroup.AppendChild(innerGroup);

        // Create the innermost shape (a rectangle) that will be placed inside the inner group.
        Shape innermostShape = new Shape(doc, ShapeType.Rectangle);
        innermostShape.Width = 100;
        innermostShape.Height = 100;
        innermostShape.Left = 10;
        innermostShape.Top = 10;
        innermostShape.Fill.Color = Color.Yellow;
        innermostShape.Stroke.Color = Color.Red;

        // Add the innermost shape to the inner group.
        innerGroup.AppendChild(innermostShape);

        // Save the document to a DOCX file.
        doc.Save("NestedShapes.docx");
    }
}
