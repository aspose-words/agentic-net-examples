using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddWatermarkExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Configure the appearance of the text watermark.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",          // Font family for the watermark text.
            FontSize = 36,                 // Font size.
            Color = Color.Gray,            // Text color.
            Layout = WatermarkLayout.Diagonal, // Diagonal layout (315° rotation).
            IsSemitrasparent = true        // Semi‑transparent (default) for a subtle effect.
        };

        // Add the text watermark to every page of the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document in DOCX format.
        doc.Save("WatermarkedDocument.docx");
    }
}
