// ALL ATTEMPTS FAILED. Below is the last generated code.

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

        // Insert a drawing canvas shape with the desired size (width, height in points).
        // ShapeType.DrawingCanvas represents a canvas that can contain other drawing objects.
        builder.InsertShape(ShapeType.DrawingCanvas, 400, 300);

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvas.docx");
    }
}
