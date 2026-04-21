using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Define paths for the sample document, image, and output document.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string docPath = Path.Combine(dataDir, "Sample.docx");
        string imagePath = Path.Combine(dataDir, "Watermark.png");
        string outputPath = Path.Combine(dataDir, "Sample_With_Image_Watermark.docx");

        // -----------------------------------------------------------------
        // Create a simple source document if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(docPath))
        {
            Document blankDoc = new Document();
            // Add a paragraph with some text so the document is not empty.
            blankDoc.FirstSection.Body.FirstParagraph.AppendChild(new Run(blankDoc, "This is a sample document."));
            blankDoc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // Create a deterministic PNG image for the watermark.
        // The image is a 1x1 pixel transparent PNG encoded in Base64.
        // -----------------------------------------------------------------
        if (!File.Exists(imagePath))
        {
            // Base64 for a 1x1 transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // -----------------------------------------------------------------
        // Load the existing document from the file system.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);

        // -----------------------------------------------------------------
        // Add the image watermark using the Document.Watermark API.
        // -----------------------------------------------------------------
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            // Scale factor (5 means 5 times the original size). Adjust as needed.
            Scale = 5,
            // Set washout to false to keep original colors.
            IsWashout = false
        };

        // Apply the watermark from the image file.
        doc.Watermark.SetImage(imagePath, options);

        // -----------------------------------------------------------------
        // Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
