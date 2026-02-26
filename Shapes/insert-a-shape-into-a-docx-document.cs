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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape with a width of 100 points and a height of 50 points.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Set visual properties of the shape (optional).
        shape.FillColor = Color.LightBlue;          // Fill the shape with a light blue color.
        shape.Stroke.Color = Color.DarkBlue;        // Outline color.
        shape.StrokeWeight = 1.5;                   // Outline thickness in points.

        // Save the document in DOCX format.
        doc.Save("ShapeInsertion.docx");
    }
}
