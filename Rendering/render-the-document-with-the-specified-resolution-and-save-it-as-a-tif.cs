using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class RenderToTiff
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Create ImageSaveOptions for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the resolution (both horizontal and vertical) in DPI.
        options.Resolution = 300;

        // Choose a compression method for the TIFF file.
        options.TiffCompression = TiffCompression.Lzw;

        // Save the rendered document as a TIFF image using the specified options.
        doc.Save("output.tiff", options);
    }
}
