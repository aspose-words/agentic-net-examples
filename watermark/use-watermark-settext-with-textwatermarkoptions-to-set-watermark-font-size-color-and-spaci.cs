using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Configure watermark options: font size, color, and layout (spacing).
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Blue,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("Confidential", options);

        // Save the document to the local file system.
        string outputPath = "Watermarked.docx";
        doc.Save(outputPath);
    }
}
