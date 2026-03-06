using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class AddWatermarkToDotm
{
    static void Main()
    {
        // Paths to the input DOTM template and the output document.
        string dataDir = @"C:\Data\";
        string inputPath = Path.Combine(dataDir, "Template.dotm");
        string outputPath = Path.Combine(dataDir, "TemplateWithWatermark.dotm");

        // Load the DOTM document.
        Document doc = new Document(inputPath);

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

        // Save the document with the watermark.
        doc.Save(outputPath);
    }
}
