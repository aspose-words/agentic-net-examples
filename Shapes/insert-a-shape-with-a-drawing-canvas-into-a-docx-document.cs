using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDrawingCanvasExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a drawing canvas – a GroupShape that can contain other shapes.
        GroupShape drawingCanvas = new GroupShape(doc);
        // Set the size and position of the canvas (in points).
        drawingCanvas.Bounds = new RectangleF(0, 0, 300, 200);
        // Optional: make the canvas floating and position it on the page.
        drawingCanvas.WrapType = WrapType.None;
        drawingCanvas.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        drawingCanvas.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        drawingCanvas.HorizontalAlignment = HorizontalAlignment.Center;
        drawingCanvas.VerticalAlignment = VerticalAlignment.Center;

        // Insert the canvas into the document.
        builder.InsertNode(drawingCanvas);

        // Create a shape (e.g., a rectangle) to place inside the drawing canvas.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        innerShape.Width = 150;
        innerShape.Height = 80;
        innerShape.Left = 20;   // Position relative to the canvas.
        innerShape.Top = 20;
        innerShape.Fill.ForeColor = Color.LightBlue;
        innerShape.Stroke.Color = Color.DarkBlue;
        innerShape.Stroke.Weight = 1.5;

        // Add the shape to the canvas.
        drawingCanvas.AppendChild(innerShape);

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvas.docx");
    }
}
