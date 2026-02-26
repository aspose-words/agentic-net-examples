using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace WatermarkPdfExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Configure text watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply the text watermark to every page of the document.
            doc.Watermark.SetText("Confidential", options);

            // Save the document as a PDF file.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            doc.Save("OutputWithWatermark.pdf", pdfOptions);
        }
    }
}
