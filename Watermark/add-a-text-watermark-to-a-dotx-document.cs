using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Configure watermark appearance.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Save the document with the watermark applied.
        doc.Save("Template_With_Watermark.dotx");
    }
}
