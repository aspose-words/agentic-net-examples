using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDrawingCanvas
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a drawing canvas shape. The canvas can contain other drawing objects.
        // Width and height are specified in points (1 point = 1/72 inch).
        // NOTE: In older Aspose.Words versions the enum value is not available; use Group as a fallback.
        Shape drawingCanvas = builder.InsertShape(ShapeType.Group, 400, 300);

        // Optional: set the canvas to be a floating shape positioned in the page center.
        drawingCanvas.WrapType = WrapType.None;                     // No text wrapping.
        drawingCanvas.BehindText = true;                           // Place behind text.
        drawingCanvas.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        drawingCanvas.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        drawingCanvas.HorizontalAlignment = HorizontalAlignment.Center;
        drawingCanvas.VerticalAlignment = VerticalAlignment.Center;

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvas.docx");
    }
}
