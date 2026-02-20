using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class AddWatermarkToDotm
{
    static void Main()
    {
        // Path to the folder that contains the DOTM template.
        string dataDir = @"C:\Data\";

        // Load the existing DOTM document.
        Document doc = new Document(dataDir + "Template.dotm");

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

        // Save the document back as a DOTM file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Dotm);
        doc.Save(dataDir + "Template_Watermarked.dotm", saveOptions);
    }
}
