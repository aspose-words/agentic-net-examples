using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define watermark text and optional formatting.
        string watermarkText = "Confidential";

        // Optional: customize appearance using TextWatermarkOptions.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText(watermarkText, options);

        // Save the document as a DOCX file.
        doc.Save("WatermarkedDocument.docx");
    }
}
