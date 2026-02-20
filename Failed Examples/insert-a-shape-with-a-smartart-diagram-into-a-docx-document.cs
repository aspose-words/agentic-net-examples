// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.SmartArt; // Namespace for SmartArt types

class InsertSmartArtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a SmartArt diagram (Basic Process layout) with specified size.
        // The InsertSmartArt method returns a Shape that contains the SmartArt.
        Shape smartArtShape = builder.InsertSmartArt(SmartArtLayoutType.BasicProcess, 400, 300);

        // Configure the shape's layout and wrapping.
        smartArtShape.WrapType = WrapType.None;          // Floating shape.
        smartArtShape.BehindText = false;               // In front of text.
        smartArtShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        smartArtShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        smartArtShape.HorizontalAlignment = HorizontalAlignment.Center;
        smartArtShape.VerticalAlignment = VerticalAlignment.Center;

        // Save the document to a DOCX file.
        doc.Save("SmartArtDiagram.docx");
    }
}
