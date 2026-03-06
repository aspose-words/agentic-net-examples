using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToImage
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\Sample.doc";

        // Path where the rendered image will be saved.
        string outputPath = @"C:\Docs\Sample.png";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Create image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Optional: render only the first page (zero‑based index).
        options.PageSet = new PageSet(0);

        // Save the document as an image.
        doc.Save(outputPath, options);
    }
}
