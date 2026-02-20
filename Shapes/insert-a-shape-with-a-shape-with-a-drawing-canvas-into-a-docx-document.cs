using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will serve as the drawing canvas.
        GroupShape canvas = new GroupShape(doc);
        canvas.Width = 300;   // Width in points.
        canvas.Height = 200;  // Height in points.
        canvas.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        canvas.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        canvas.WrapType = WrapType.None;
        canvas.BehindText = true;
        canvas.HorizontalAlignment = HorizontalAlignment.Center;
        canvas.VerticalAlignment = VerticalAlignment.Center;

        // Insert the canvas into the document.
        builder.InsertNode(canvas);

        // Create a rectangle shape to place inside the canvas.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 50;   // Position relative to the canvas.
        rect.Top = 30;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;
        rect.StrokeWeight = 1.5;

        // Add the rectangle to the canvas.
        canvas.AppendChild(rect);

        // Save the document.
        doc.Save("ShapeWithCanvas.docx");
    }
}
