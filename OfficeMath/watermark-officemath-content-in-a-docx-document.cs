using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document that contains OfficeMath content.
        Document doc = new Document("Input.docx");

        // Configure the text watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
