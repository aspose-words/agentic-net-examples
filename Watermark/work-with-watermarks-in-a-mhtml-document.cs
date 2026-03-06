using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkMhtmlExample
{
    static void Main()
    {
        // Load an existing MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Create text watermark options (optional customization).
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Save the document back to MHTML format.
        doc.Save("OutputDocument.mhtml");
    }
}
