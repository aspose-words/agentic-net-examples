using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline shape (50x50 points) of type DiagonalCornersRounded.
        Shape inlineShape = builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);
        // Example formatting: set the outline color.
        inlineShape.Stroke.Color = Color.Blue;

        // Insert a floating shape (TopCornersRounded) positioned 100 points from the page edges.
        Shape floatingShape = builder.InsertShape(
            ShapeType.TopCornersRounded,
            RelativeHorizontalPosition.Page, 100,   // Horizontal position relative to page.
            RelativeVerticalPosition.Page, 100,     // Vertical position relative to page.
            80, 80,                                 // Width and height in points.
            WrapType.None);                         // No text wrapping.
        floatingShape.Stroke.Color = Color.Red;

        // Save the document using ISO 29500 Transitional compliance so that DML shapes are preserved.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("ShapesInsertion.docx", saveOptions);
    }
}
