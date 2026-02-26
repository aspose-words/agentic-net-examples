using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark appearance.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the document as RTF using RtfSaveOptions.
        RtfSaveOptions rtfOptions = new RtfSaveOptions();
        doc.Save("WatermarkedDocument.rtf", rtfOptions);
    }
}
