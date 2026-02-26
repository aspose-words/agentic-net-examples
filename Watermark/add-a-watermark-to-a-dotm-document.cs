using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddWatermarkToDotm
{
    static void Main()
    {
        // Path to the folder that contains the DOTM template.
        string dataDir = @"C:\Data\";

        // Load the existing DOTM document.
        Document doc = new Document(Path.Combine(dataDir, "Template.dotm"));

        // Configure text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the document with the watermark applied.
        doc.Save(Path.Combine(dataDir, "Template_Watermarked.dotm"));
    }
}
