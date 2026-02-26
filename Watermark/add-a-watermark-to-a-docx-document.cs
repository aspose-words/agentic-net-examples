using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Path where the output document will be saved.
        string outputPath = "Watermarked.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Define options for the text watermark.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add the text watermark with the defined options.
        doc.Watermark.SetText("Confidential", options);

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
