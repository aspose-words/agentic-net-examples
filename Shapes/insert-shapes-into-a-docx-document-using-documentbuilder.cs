using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape (width: 100 points, height: 50 points).
        Shape inlineShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        inlineShape.FillColor = Color.LightBlue;
        inlineShape.StrokeColor = Color.DarkBlue;

        // Insert a floating star shape positioned 100 points from the left and top of the page.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Star,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            80, 80,
            WrapType.None);
        floatingShape.FillColor = Color.Yellow;
        floatingShape.StrokeColor = Color.Orange;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(inlineShape, floatingShape);
        // Optionally reposition the group.
        group.Left = 150;
        group.Top = 200;

        // Save the document with OOXML Transitional compliance to preserve DML shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("ShapesInserted.docx", saveOptions);
    }
}
