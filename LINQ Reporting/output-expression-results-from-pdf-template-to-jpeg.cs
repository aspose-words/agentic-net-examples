using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OutputExpressionResultsFromPdfTemplateToJpeg
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the PDF template that contains expression fields (MERGEFIELD, IF, etc.).
            Document doc = new Document("Template.pdf");

            // Ensure that all fields are evaluated and the results are reflected in the document.
            doc.UpdateFields();

            // Configure image save options for JPEG output.
            // The JpegQuality property controls the compression level (0‑100).
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80,               // Moderate compression with good visual quality.
                PageSet = new PageSet(0)        // Render the first page (zero‑based index). Adjust as needed.
            };

            // Save the rendered page(s) as a JPEG image.
            doc.Save("Result.jpg", jpegOptions);
        }
    }
}
