using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document as a plain‑text (TXT) file.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("Output.txt", saveOptions);
    }
}
