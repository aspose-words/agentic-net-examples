using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the modified document as a DOTX file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Dotx);
        doc.Save("WatermarkedTemplate.dotx", saveOptions);
    }
}
