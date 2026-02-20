using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace InsertGroupShapeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a GroupShape that will hold other shapes.
            GroupShape group = new GroupShape(doc);

            // Set the size and position of the group shape (in points).
            group.Bounds = new RectangleF(0, 0, 200, 200);

            // Insert the group shape into the current paragraph.
            builder.CurrentParagraph.AppendChild(group);

            // Create a rectangle shape to place inside the group.
            Shape rect = new Shape(doc, ShapeType.Rectangle);
            rect.Width = 100;
            rect.Height = 100;
            rect.Left = 0;
            rect.Top = 0;
            // Optional: set a simple fill colour so the shape is visible.
            rect.FillColor = Color.LightBlue;

            // Add the rectangle to the group shape.
            group.AppendChild(rect);

            // Create a second shape (an ellipse) inside the same group.
            Shape ellipse = new Shape(doc, ShapeType.Ellipse);
            ellipse.Width = 80;
            ellipse.Height = 80;
            ellipse.Left = 100;
            ellipse.Top = 100;
            ellipse.FillColor = Color.LightCoral;

            // Add the ellipse to the group shape.
            group.AppendChild(ellipse);

            // Save the document containing the group shape.
            doc.Save("GroupShape.docx");
        }
    }
}
