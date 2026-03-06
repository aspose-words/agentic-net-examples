// ALL ATTEMPTS FAILED. Below is the last generated code.

using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a drawing canvas (ShapeType.Canvas) with the desired size (width x height in points).
        // Here we use 400 points wide and 300 points high as an example.
        Shape canvas = builder.InsertShape(ShapeType.Canvas, 400, 300);

        // Optional: set a visible border so the canvas can be seen in the document.
        canvas.LineFormat.Color = Color.Black;
        canvas.LineFormat.Width = 1.0;

        // Save the document to a DOCX file.
        doc.Save("DrawingCanvas.docx");
    }
}
