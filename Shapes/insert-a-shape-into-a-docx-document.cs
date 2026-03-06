using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape with width=100 points and height=50 points.
        // ShapeType.Rectangle is a primitive shape supported by InsertShape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Optional: customize the shape (e.g., set fill color, stroke).
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;
        shape.StrokeWeight = 1.5;

        // Save the document to a DOCX file.
        doc.Save("ShapeInsertion.docx");
    }
}
