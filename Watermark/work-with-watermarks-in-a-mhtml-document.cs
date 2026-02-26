using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class WatermarkMhtmlExample
{
    static void Main()
    {
        // Path to the source MHTML document.
        string inputPath = @"C:\Docs\source.mhtml";

        // Path where the watermarked MHTML document will be saved.
        string outputPath = @"C:\Docs\watermarked.mhtml";

        // Load the MHTML document.
        Document doc = new Document(inputPath);

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Black,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to every page of the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document back to MHTML format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
