using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add some sample text to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will be saved as PDF with a watermark.");

        // Configure watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document as a PDF file.
        doc.Save("DocumentWithWatermark.pdf", SaveFormat.Pdf);
    }
}
