using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape (inline) with width 150 points and height 100 points.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rectangle.Left = 50;   // Optional positioning.
        rectangle.Top = 50;
        rectangle.Stroke.Color = Color.Blue;

        // Insert an ellipse shape (inline) with width 120 points and height 120 points.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        ellipse.Left = 250;    // Optional positioning.
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the rectangle and ellipse into a single GroupShape.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);
        group.Left = 20;       // Position the group.
        group.Top = 20;

        // Save the document as DOCX with DML compliance to preserve shape types.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("ChildShapes.docx", saveOptions);
    }
}
