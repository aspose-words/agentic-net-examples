using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class WatermarkDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a text watermark with custom formatting.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the document with the watermark.
        string outPath = "Watermarked.docx";
        doc.Save(outPath, SaveFormat.Docx);

        // Load the saved document.
        Document loadedDoc = new Document(outPath);

        // If a text watermark exists, remove it.
        if (loadedDoc.Watermark.Type == WatermarkType.Text)
        {
            loadedDoc.Watermark.Remove();
        }

        // Save the document after removing the watermark.
        string finalPath = "WatermarkRemoved.docx";
        loadedDoc.Save(finalPath, SaveFormat.Docx);
    }
}
