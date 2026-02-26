using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsTiff
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting TIFF file will be saved.
        string outputPath = @"C:\Docs\ResultImage.tiff";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Remove headers and footers from the first (and only) section.
        // This clears the content of the header/footer objects, effectively omitting them from the rendered pages.
        doc.FirstSection.ClearHeadersFooters();

        // Configure image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set compression type (default is Lzw). Adjust as needed.
            // TiffCompression = TiffCompression.None,
            // Optional: set resolution (dpi) for higher quality output.
            // Resolution = 300
        };

        // Save the document as a TIFF image. Each page will be rendered to a separate frame in the TIFF file.
        doc.Save(outputPath, saveOptions);
    }
}
