using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Path to the output document.
        string outputPath = "Watermarked.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Configure the text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",          // Custom font family.
            FontSize = 36,                 // Font size in points.
            Color = Color.Black,           // Font color.
            Layout = WatermarkLayout.Diagonal, // Diagonal layout.
            IsSemitrasparent = false      // Opaque watermark.
        };

        // Add the text watermark with the specified options.
        doc.Watermark.SetText("Aspose Watermark", options);

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
