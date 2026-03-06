using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDrawingCanvasExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a group shape that will act as a drawing canvas (container for other shapes).
        // Width and height are specified in points.
        Shape canvas = builder.InsertShape(ShapeType.Group, 300, 200);
        // Make the canvas floating and without text wrapping.
        canvas.WrapType = WrapType.None;
        // Optional: give the canvas a transparent fill so it does not obscure underlying text.
        canvas.FillColor = Color.Transparent;

        // Create a rectangle shape that will be placed inside the drawing canvas.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 100;          // Width in points.
        rectangle.Height = 50;          // Height in points.
        rectangle.Left = 10;            // Position relative to the canvas.
        rectangle.Top = 10;
        rectangle.FillColor = Color.LightBlue;
        rectangle.Stroke.Color = Color.DarkBlue;

        // Add the rectangle as a child of the canvas.
        canvas.AppendChild(rectangle);

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvas.docx");
    }
}
