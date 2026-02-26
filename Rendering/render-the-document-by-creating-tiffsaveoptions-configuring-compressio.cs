using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderToTiff
{
    static void Main()
    {
        // Load an existing Word document.
        // Replace with the actual path to your source document.
        string inputPath = @"C:\Docs\InputDocument.docx";
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the desired TIFF compression (e.g., LZW).
        tiffOptions.TiffCompression = TiffCompression.Lzw;

        // Set the resolution (dots per inch) for the rendered image.
        tiffOptions.Resolution = 300; // 300 DPI

        // Define the page range to render (pages 1 through 2, zero‑based indexing).
        // PageRange(start, end) includes both start and end pages.
        tiffOptions.PageSet = new PageSet(new PageRange(0, 1));

        // Save the document as a TIFF image using the configured options.
        // Replace with the desired output path.
        string outputPath = @"C:\Docs\RenderedOutput.tiff";
        doc.Save(outputPath, tiffOptions);
    }
}
