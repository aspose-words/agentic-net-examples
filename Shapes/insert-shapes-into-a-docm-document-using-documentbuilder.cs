using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertShapesIntoDocm
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Insert a floating shape (rounded rectangle) with no text wrapping.
        // -----------------------------------------------------------------
        // Parameters: shape type, horizontal position, left offset,
        // vertical position, top offset, width, height, wrap type.
        Shape floatingShape = builder.InsertShape(
            ShapeType.TopCornersRounded,
            RelativeHorizontalPosition.Page, 100,   // 100 points from the left of the page
            RelativeVerticalPosition.Page, 100,     // 100 points from the top of the page
            100, 50,                                // width = 100 points, height = 50 points
            WrapType.None);                         // shape floats behind text

        // Optional: give the shape a name and a fill colour.
        floatingShape.Name = "FloatingRoundedRect";
        floatingShape.Fill.ForeColor = System.Drawing.Color.LightBlue;

        // -----------------------------------------------------------------
        // Insert an inline shape (diagonal rounded rectangle) that behaves like a character.
        // -----------------------------------------------------------------
        Shape inlineShape = builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);
        inlineShape.Name = "InlineDiagonalRounded";
        inlineShape.Fill.ForeColor = System.Drawing.Color.LightGreen;

        // -----------------------------------------------------------------
        // Save the document as a DOCM file.
        // Use ISO/IEC 29500:2008 Transitional compliance so that non‑primitive shapes
        // are stored using DrawingML (required for the rounded rectangle shape).
        // -----------------------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        doc.Save("ShapeInsertion.docm", saveOptions);
    }
}
