using System;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffCompressionExample
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"ArtifactsDir\SourceDocument.docx";

        // Path where the rendered TIFF image will be saved.
        string outputPath = @"ArtifactsDir\RenderedDocument.tiff";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Specify the desired compression scheme.
        // Available values: TiffCompression.None, Rle, Lzw, Ccitt3, Ccitt4.
        options.TiffCompression = TiffCompression.Lzw; // Example: LZW compression.

        // Render the document pages to TIFF using the configured options.
        doc.Save(outputPath, options);
    }
}
