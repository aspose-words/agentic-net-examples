using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;   // For Color

class WatermarkToPdf
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Add a text watermark to the document.
        // The watermark will be diagonal, black, 36‑point Arial font.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false   // Make the watermark fully opaque.
        };
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document directly to PDF format.
        // The Save method automatically determines the format from the file extension,
        // but we explicitly specify SaveFormat.Pdf for clarity.
        doc.Save("WatermarkedDocument.pdf", SaveFormat.Pdf);
    }
}
