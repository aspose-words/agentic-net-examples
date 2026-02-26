using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the image that will be used as a watermark.
        string imagePath = @"C:\Images\Logo.jpg";

        // Folder where the Markdown file will be saved.
        string outputFolder = @"C:\Output";

        // Ensure the output folder exists.
        System.IO.Directory.CreateDirectory(outputFolder);

        // Create a new empty document.
        Document doc = new Document();

        // Optional: add some sample text so the Markdown file is not empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with an image watermark.");

        // Configure image watermark options.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor for the watermark image.
            IsWashout = false   // Disable washout effect.
        };

        // Add the image watermark to the document.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Set up Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Specify a folder where images extracted from the document will be saved.
            ImagesFolder = System.IO.Path.Combine(outputFolder, "Images")
        };

        // Save the document as a Markdown file.
        string markdownPath = System.IO.Path.Combine(outputFolder, "DocumentWithWatermark.md");
        doc.Save(markdownPath, saveOptions);
    }
}
