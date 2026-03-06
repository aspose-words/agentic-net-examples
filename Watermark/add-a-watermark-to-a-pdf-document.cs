using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Define watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to every page.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the result as a PDF file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("Output.pdf", pdfOptions);
    }
}
