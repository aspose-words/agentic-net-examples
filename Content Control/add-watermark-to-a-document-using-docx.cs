using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

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
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to the document.
        doc.Watermark.SetText("Aspose Watermark", options);

        // Save the document as a DOCX file.
        doc.Save("WatermarkedDocument.docx");
    }
}
