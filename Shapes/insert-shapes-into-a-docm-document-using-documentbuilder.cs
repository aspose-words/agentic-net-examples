using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank DOCM document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape (width: 100 points, height: 50 points).
        Shape inlineShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        inlineShape.Fill.ForeColor = Color.LightBlue;
        inlineShape.Stroke.Color = Color.DarkBlue;

        // Insert a floating star shape positioned relative to the page.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Star,
            RelativeHorizontalPosition.Page, 150,   // 150 points from the left edge of the page
            RelativeVerticalPosition.Page, 200,     // 200 points from the top edge of the page
            80, 80,                                 // width and height in points
            WrapType.Square);                       // text wrapping style
        floatingShape.Fill.ForeColor = Color.Yellow;
        floatingShape.Stroke.Color = Color.Orange;

        // Save the document as a macro-enabled DOCM file with ISO/IEC 29500 Transitional compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        doc.Save("ShapesInserted.docm", saveOptions);
    }
}
