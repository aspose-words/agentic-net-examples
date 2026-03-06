using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Configure image watermark options.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Example scale factor.
            IsWashout = false   // Disable washout effect.
        };

        // Path to the image that will be used as a watermark.
        string imagePath = @"C:\Images\Logo.jpg";

        // Add the image watermark to the document.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Prepare Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Save extracted images to a subfolder named "Images".
            ImagesFolder = Path.Combine(Directory.GetCurrentDirectory(), "Images")
            // Uncomment the following line to embed images as Base64 instead of separate files.
            // ExportImagesAsBase64 = true
        };

        // Save the document as a Markdown file.
        doc.Save("Output.md", saveOptions);
    }
}
