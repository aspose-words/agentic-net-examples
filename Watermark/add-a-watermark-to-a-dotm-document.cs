using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddWatermarkToDotm
{
    static void Main()
    {
        // Path to the folder containing the DOTM template.
        string dataDir = @"C:\Data\";

        // Load the existing DOTM document.
        Document doc = new Document(dataDir + "Template.dotm");

        // Configure text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Save the document with the watermark applied.
        doc.Save(dataDir + "Template_Watermarked.dotm");
    }
}
