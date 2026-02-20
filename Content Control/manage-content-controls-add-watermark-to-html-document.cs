using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Configure text watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("Confidential", watermarkOptions);

        // Save the document back to HTML format.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        doc.Save("output.html", saveOptions);
    }
}
