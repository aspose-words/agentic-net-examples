using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape (inline) with width=100pt and height=50pt.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Set a fill color for the rectangle.
        rectangle.FillColor = Color.LightBlue;

        // Insert an ellipse shape (inline) with width=80pt and height=80pt.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 80, 80);
        // Set a fill color for the ellipse.
        ellipse.FillColor = Color.LightCoral;

        // Save the document to a DOCX file.
        doc.Save("ChildShapes.docx");
    }
}
