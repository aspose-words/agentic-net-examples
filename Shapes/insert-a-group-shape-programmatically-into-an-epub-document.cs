using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoEpub
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two individual shapes that will later be grouped.
        // Shape 1: Rectangle, 200x250 points.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;   // Position relative to the page.
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        // Shape 2: Ellipse, 150x200 points.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. InsertGroupShape automatically calculates
        // the position and size of the resulting GroupShape.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Create a third shape (triangle) that will be added to the existing group.
        // ShapeType is read‑only, so we must instantiate a new Shape with the desired type.
        Shape triangle = new Shape(doc, ShapeType.Triangle);
        triangle.Width = 100;   // Adjust size as needed.
        triangle.Height = 100;
        triangle.Stroke.Color = Color.Blue;

        // Add the triangle to the previously created group.
        group.AppendChild(triangle);

        // Save the document as an EPUB file.
        doc.Save("GroupShape.epub", SaveFormat.Epub);
    }
}
