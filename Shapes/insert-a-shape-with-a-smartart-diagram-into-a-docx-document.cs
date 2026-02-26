using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertSmartArtShape
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating shape that will hold the SmartArt diagram.
        // The shape type can be any that supports SmartArt (e.g., Group).
        Shape smartArtShape = builder.InsertShape(ShapeType.Group, 300, 200);
        // Position the shape on the page (optional).
        smartArtShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        smartArtShape.RelativeVerticalPosition   = RelativeVerticalPosition.Page;
        smartArtShape.Left  = 100; // points from the left edge of the page
        smartArtShape.Top   = 100; // points from the top edge of the page
        smartArtShape.WrapType = WrapType.None;

        // If the shape already contains a SmartArt object, update its drawing.
        // (When inserting a new SmartArt this call ensures the pre‑rendered drawing is generated.)
        if (smartArtShape.HasSmartArt)
        {
            smartArtShape.UpdateSmartArtDrawing();
        }

        // Save the document with OOXML compliance that supports DML shapes (required for SmartArt).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("SmartArtShape.docx", saveOptions);
    }
}
