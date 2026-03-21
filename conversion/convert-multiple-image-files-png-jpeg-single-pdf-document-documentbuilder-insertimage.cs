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
        // Use folders relative to the executable location.
        string baseDir = AppContext.BaseDirectory;
        string imagesFolder = Path.Combine(baseDir, "Images");
        string outputFolder = Path.Combine(baseDir, "Result");
        Directory.CreateDirectory(imagesFolder);
        Directory.CreateDirectory(outputFolder);

        // Path for the resulting PDF document.
        string outputPdfPath = Path.Combine(outputFolder, "CombinedImages.pdf");

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get all PNG and JPEG files from the folder (non‑recursive).
        string[] imageFiles = Directory.GetFiles(imagesFolder)
            .Where(f => f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (imageFiles.Length == 0)
        {
            Console.WriteLine($"No image files found in '{imagesFolder}'. Place PNG/JPEG files there and rerun.");
            return;
        }

        // Insert each image on a separate page.
        for (int i = 0; i < imageFiles.Length; i++)
        {
            builder.InsertImage(imageFiles[i]);

            // Add a page break after each image except the last one.
            if (i < imageFiles.Length - 1)
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
        Console.WriteLine($"PDF created successfully at '{outputPdfPath}'.");
    }
}
