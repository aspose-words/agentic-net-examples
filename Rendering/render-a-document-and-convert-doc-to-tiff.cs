using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"MyDir\Document.doc";

        // Path where the resulting TIFF file will be saved.
        string outputPath = @"ArtifactsDir\Document.tiff";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Optional: set resolution (dpi) and compression if needed.
        // tiffOptions.Resolution = 300; // 300 DPI
        // tiffOptions.TiffCompression = TiffCompression.Lzw; // LZW compression

        // Save the document as a TIFF image (each page will be rendered to a separate frame).
        doc.Save(outputPath, tiffOptions);
    }
}
