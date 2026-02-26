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

        // Add a simple text watermark with default settings.
        doc.Watermark.SetText("Confidential");

        // Configure custom options for a styled text watermark.
        TextWatermarkOptions textOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the styled text watermark.
        doc.Watermark.SetText("Do Not Distribute", textOptions);

        // Save the document containing the watermark.
        doc.Save("Watermarked.docx");

        // Load the previously saved document.
        Document loadedDoc = new Document("Watermarked.docx");

        // Remove the watermark if it is a text watermark.
        if (loadedDoc.Watermark.Type == WatermarkType.Text)
        {
            loadedDoc.Watermark.Remove();
        }

        // Configure options for an image watermark.
        ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
        {
            Scale = 3,          // Increase the size of the image.
            IsWashout = true    // Make the image semi‑transparent.
        };

        // Add an image watermark from a file using the configured options.
        loadedDoc.Watermark.SetImage("logo.png", imageOptions);

        // Save the document with the image watermark.
        loadedDoc.Save("ImageWatermarked.docx");
    }
}
