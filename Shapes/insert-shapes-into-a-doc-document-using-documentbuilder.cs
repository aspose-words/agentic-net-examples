using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        // Parameters: shape type, horizontal position, left offset,
        // vertical position, top offset, width, height, wrap type.
        Shape floatingRect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            150, 80,
            WrapType.None);

        // Make the shape appear behind the text.
        floatingRect.BehindText = true;

        // Insert an inline ellipse shape (width 100, height 50).
        // Inline shapes are placed directly in the paragraph.
        Shape inlineEllipse = builder.InsertShape(ShapeType.Ellipse, 100, 50);

        // Set OOXML compliance to allow saving of non‑primitive shapes.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document to a file.
        doc.Save("Shapes.docx", saveOptions);
    }
}
