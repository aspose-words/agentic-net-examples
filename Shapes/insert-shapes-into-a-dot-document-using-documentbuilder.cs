using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape with a width of 100 points and a height of 50 points.
        Shape inlineShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Set visual properties for the inline shape.
        inlineShape.Stroke.Color = Color.Blue;
        inlineShape.Fill.ForeColor = Color.LightGray;

        // Insert a floating ellipse shape positioned 100 points from the left and top of the page.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            150, // width in points
            100, // height in points
            WrapType.None);
        // Set visual properties for the floating shape.
        floatingShape.Stroke.Color = Color.Red;
        floatingShape.Fill.ForeColor = Color.Yellow;

        // Save the document to a file.
        doc.Save("ShapesDocument.docx");
    }
}
