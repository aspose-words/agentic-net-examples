using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Configure options for a text watermark.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Save the document that now contains the watermark.
        doc.Save("Watermarked.docx");

        // Load the previously saved document.
        Document loadedDoc = new Document("Watermarked.docx");

        // If a text watermark exists, remove it.
        if (loadedDoc.Watermark.Type == WatermarkType.Text)
        {
            loadedDoc.Watermark.Remove();
        }

        // Save the document after the watermark has been removed.
        loadedDoc.Save("WatermarkedRemoved.docx");
    }
}
