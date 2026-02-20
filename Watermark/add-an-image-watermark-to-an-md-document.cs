using System.Drawing;                     // For Image class
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class AddImageWatermarkToMarkdown
{
    static void Main()
    {
        // Path to the source markdown file.
        const string inputPath = @"C:\Docs\source.md";

        // Path to the image that will be used as a watermark.
        const string imagePath = @"C:\Images\Logo.jpg";

        // Path where the watermarked markdown will be saved.
        const string outputPath = @"C:\Docs\watermarked.md";

        // Load the markdown document.
        var loadOptions = new MarkdownLoadOptions();               // Load options for markdown
        Document doc = new Document(inputPath, loadOptions);

        // Configure the appearance of the image watermark.
        var imageWatermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor (5 times the original size)
            IsWashout = false  // Disable washout effect for a solid image
        };

        // Add the image watermark to the document.
        // The overload that accepts a file path and options is used.
        doc.Watermark.SetImage(imagePath, imageWatermarkOptions);

        // Save the document back to markdown format.
        var saveOptions = new MarkdownSaveOptions();               // Default markdown save options
        doc.Save(outputPath, saveOptions);
    }
}
