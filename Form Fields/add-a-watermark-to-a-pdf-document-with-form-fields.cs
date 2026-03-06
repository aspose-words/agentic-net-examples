using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddWatermarkToPdf
{
    static void Main()
    {
        // Paths to the source PDF (which may contain form fields) and the output PDF.
        string inputPath = @"C:\Docs\Input.pdf";
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the PDF document.
        Document doc = new Document(inputPath);

        // Configure watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Add a text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the modified document as PDF, preserving any existing form fields.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
