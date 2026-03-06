using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithShape
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the outer shape (a rectangle) but do NOT insert it yet.
        Shape outerShape = new Shape(builder.Document, ShapeType.Rectangle);
        outerShape.Width = 200;   // Width in points.
        outerShape.Height = 100;  // Height in points.
        outerShape.Stroke.Color = Color.Blue; // Outline color.

        // Create the inner shape (an ellipse) that will be placed inside the outer shape.
        Shape innerShape = new Shape(builder.Document, ShapeType.Ellipse);
        innerShape.Width = 50;
        innerShape.Height = 50;
        innerShape.Stroke.Color = Color.Red;

        // Group the two shapes together. The InsertGroupShape method inserts a GroupShape
        // containing the provided shapes at the current cursor position.
        GroupShape group = builder.InsertGroupShape(outerShape, innerShape);

        // Optionally adjust the position of the group shape.
        group.Left = 100;   // Distance from the left edge of the page.
        group.Top = 100;    // Distance from the top edge of the page.

        // Save the document in DOCX format.
        doc.Save("ShapeWithShape.docx");
    }
}
