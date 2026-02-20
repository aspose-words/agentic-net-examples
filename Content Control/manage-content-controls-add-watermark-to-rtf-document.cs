using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class AddWatermarkToRtf
{
    static void Main()
    {
        // Load an existing RTF document.
        Document doc = new Document("input.rtf");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document as RTF using RtfSaveOptions.
        RtfSaveOptions saveOptions = new RtfSaveOptions();
        doc.Save("output.rtf", saveOptions);
    }
}
