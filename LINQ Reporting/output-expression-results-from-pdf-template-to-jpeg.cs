using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OutputExpressionResultsFromPdfTemplateToJpeg
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the PDF template that contains fields (expressions) to be evaluated.
            Document doc = new Document("Template.pdf");

            // Configure image save options for JPEG output.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Update fields (MERGEFIELD, REF, etc.) before rendering.
                UpdateFields = true,

                // Set JPEG compression quality (0‑100). Adjust as needed.
                JpegQuality = 90,

                // Optional: render each page separately. Comment out if a single image of the first page is required.
                // PageSet = new PageSet(0) // renders only the first page.
            };

            // Save the entire document as a JPEG image (one image per page).
            doc.Save("Result.jpg", jpegOptions);

            // If you need separate JPEG files for each page, uncomment the loop below.
            /*
            for (int i = 0; i < doc.PageCount; i++)
            {
                jpegOptions.PageSet = new PageSet(i); // Render page i only.
                string outputPath = $"Result_Page_{i + 1}.jpg";
                doc.Save(outputPath, jpegOptions);
            }
            */
        }
    }
}
