using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape positioned 100 points from the left and top of the page.
        // The shape is 200 points wide and 100 points high, with no text wrapping.
        builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            200, 100,
            WrapType.None);

        // Insert an inline ellipse shape with a width and height of 100 points.
        builder.InsertShape(ShapeType.Ellipse, 100, 100);

        // Configure save options to use a newer OOXML compliance level.
        // This ensures that non‑primitive shapes are saved using DrawingML (DML).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document to a DOCX file.
        doc.Save("Shapes.docx", saveOptions);
    }
}
