using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExportGroupShapeToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two individual shapes.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;
        rect.Top = 50;
        rect.Stroke.Color = Color.Red;
        rect.FillColor = Color.LightYellow;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 100);
        ellipse.Left = 300;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Blue;
        ellipse.FillColor = Color.LightGreen;

        // Group the two shapes into a single GroupShape and insert it at the current position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust group properties (e.g., position, size) if needed.
        group.Left = 20;
        group.Top = 20;

        // Save the document as PDF. The format is inferred from the file extension.
        doc.Save("GroupShapeDocument.pdf", SaveFormat.Pdf);
    }
}
