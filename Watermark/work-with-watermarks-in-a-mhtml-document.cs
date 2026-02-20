using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class WatermarkMhtmlExample
{
    static void Main()
    {
        // Load an existing MHTML document.
        // The Document constructor automatically detects the format.
        Document doc = new Document("InputDocument.mhtml");

        // Create text watermark options.
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add a text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", textOptions);

        // Prepare save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Ensure resources (images, CSS) are embedded using CID URLs.
            ExportCidUrlsForMhtmlResources = true,
            // Keep the original document properties.
            ExportDocumentProperties = true
        };

        // Save the modified document back to MHTML.
        doc.Save("OutputDocument.mhtml", saveOptions);
    }
}
