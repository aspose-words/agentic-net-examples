using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class RenderDocumentToTiff
{
    static void Main()
    {
        // Load the source Word document.
        // Replace with the actual path to your document.
        string inputPath = @"C:\Docs\input.docx";
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the desired resolution (dots per inch) for both horizontal and vertical axes.
        options.Resolution = 300f; // 300 DPI

        // Optional: specify TIFF compression (default is Lzw).
        // options.TiffCompression = TiffCompression.Lzw;

        // Save the rendered document as a TIFF image.
        // Replace with the desired output path.
        string outputPath = @"C:\Docs\output.tiff";
        doc.Save(outputPath, options);
    }
}
