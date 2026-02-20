// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertSmartArtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a SmartArt diagram. 
        // The InsertSmartArt method creates a Shape that contains a SmartArt object.
        // Parameters: layout type, width (points), height (points).
        // Here we use a basic process layout as an example.
        Shape smartArtShape = builder.InsertSmartArt(
            Aspose.Words.SmartArt.SmartArtLayoutType.BasicProcess, 400, 300);

        // Optionally set the shape's wrapping style and position.
        smartArtShape.WrapType = WrapType.None;               // Floating shape.
        smartArtShape.BehindText = false;                    // In front of text.
        smartArtShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        smartArtShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        smartArtShape.HorizontalAlignment = HorizontalAlignment.Center;
        smartArtShape.VerticalAlignment = VerticalAlignment.Center;

        // Save the document in DOCX format.
        doc.Save("SmartArtShape.docx", SaveFormat.Docx);
    }
}
