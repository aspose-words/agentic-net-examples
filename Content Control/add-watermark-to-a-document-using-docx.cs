using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Path where the output document will be saved.
        string outputPath = @"C:\Temp\Watermarked.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Configure text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the document in DOCX format.
        doc.Save(outputPath);
    }
}
