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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a drawing canvas shape (inline) with the desired size.
        // In recent versions of Aspose.Words the drawing canvas is represented by ShapeType.Group.
        Shape canvas = builder.InsertShape(ShapeType.Group, 300, 200);
        canvas.WrapType = WrapType.None;          // Make it floating (no text wrap).
        canvas.BehindText = true;                // Place it behind the document text.

        // Create a rectangle shape that will be placed inside the canvas.
        Shape innerRectangle = new Shape(doc, ShapeType.Rectangle);
        innerRectangle.Width = 100;
        innerRectangle.Height = 50;
        innerRectangle.Left = 20;                // Position relative to the canvas origin.
        innerRectangle.Top = 20;
        innerRectangle.FillColor = Color.LightBlue;
        innerRectangle.Stroke.Color = Color.DarkBlue;

        // Append the inner shape to the canvas so it becomes a child of the drawing canvas.
        canvas.AppendChild(innerRectangle);

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvasShape.docx");
    }
}
