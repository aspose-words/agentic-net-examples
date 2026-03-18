using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImagesToPdfConverter
{
    static void Main()
    {
        // Folder that contains the source PNG and JPEG images.
        string imagesFolder = @"C:\Images";

        // Path for the resulting PDF document.
        string outputPdfPath = @"C:\Result\CombinedImages.pdf";

        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to insert content into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get all PNG and JPEG files from the folder (non‑recursive).
        string[] imageFiles = Directory.GetFiles(imagesFolder)
            .Where(f => f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        // Insert each image on a separate page.
        foreach (string imagePath in imageFiles)
        {
            // Insert the image inline at its original dimensions.
            builder.InsertImage(imagePath);

            // Add a page break after each image except the last one.
            if (imagePath != imageFiles.Last())
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Optional: configure PDF save options (e.g., JPEG compression for all images).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 90 // Adjust quality as needed (0‑100).
        };

        // Save the document as a PDF file using the specified options.
        doc.Save(outputPdfPath, pdfOptions);
    }
}
