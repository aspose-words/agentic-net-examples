using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing RTF document.
        Document doc = new Document("input.rtf");

        // Define watermark appearance (optional).
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add a text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the modified document back to RTF format.
        doc.Save("output.rtf");
    }
}
