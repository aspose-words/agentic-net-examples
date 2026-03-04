using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Define watermark appearance.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to every page.
        doc.Watermark.SetText("Confidential", options);

        // Save the modified document as HTML.
        doc.Save("output.html");
    }
}
