using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure text watermark options.
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

        // Save the document with the watermark.
        doc.Save("Watermarked.docx");

        // If a text watermark exists, remove it.
        if (doc.Watermark.Type == WatermarkType.Text)
        {
            doc.Watermark.Remove();

            // Save the document after removing the watermark.
            doc.Save("WatermarkRemoved.docx");
        }
    }
}
