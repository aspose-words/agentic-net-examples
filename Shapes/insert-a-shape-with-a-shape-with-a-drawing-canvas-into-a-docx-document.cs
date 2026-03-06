using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertShapeWithCanvas
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a drawing canvas – in Aspose.Words this can be represented by a GroupShape.
        GroupShape canvas = new GroupShape(doc);
        // Set the canvas size and position (in points).
        canvas.Bounds = new RectangleF(0, 0, 300, 200);
        // Define the internal coordinate system for the canvas.
        canvas.CoordSize = new Size(500, 500);

        // Create a shape that will be placed inside the canvas.
        Shape innerShape = new Shape(doc, ShapeType.Rectangle);
        innerShape.Width = 150;   // width in points
        innerShape.Height = 100;  // height in points
        innerShape.Left = 50;     // position relative to the canvas origin
        innerShape.Top = 30;
        innerShape.Stroke.Color = Color.Blue;
        innerShape.FillColor = Color.LightGray;

        // Add the shape to the canvas.
        canvas.AppendChild(innerShape);

        // Insert the canvas (group shape) into the document at the current cursor position.
        builder.InsertNode(canvas);

        // Save the document as DOCX with strict compliance to preserve the drawing canvas.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        doc.Save("ShapeWithCanvas.docx", saveOptions);
    }
}
