// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertSmartArtExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a SmartArt diagram. Choose a layout (e.g., BasicCycle) and set size.
        // The InsertSmartArt method returns the Shape that contains the SmartArt.
        Shape smartArtShape = builder.InsertSmartArt(SmartArtLayout.BasicCycle, 400, 300);

        // Optionally set a title for the SmartArt shape.
        smartArtShape.Title = "Sample SmartArt";

        // Save the document to a DOCX file.
        doc.Save("SmartArtDiagram.docx");
    }
}
