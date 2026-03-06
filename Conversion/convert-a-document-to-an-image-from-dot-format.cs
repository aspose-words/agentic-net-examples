using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToImageConverter
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        const string inputPath = @"C:\Docs\Template.dot";

        // Path where the rendered image will be saved.
        const string outputPath = @"C:\Docs\TemplateImage.png";

        // Load the DOT document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Configure image save options.
        // SaveFormat.Png renders the first page of the document as a PNG image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300
        };

        // Save the document as an image using the configured options.
        doc.Save(outputPath, options);
    }
}
