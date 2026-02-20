using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Set up options for rendering the document to a TIFF image.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Example compression: LZW (you can choose other values from TiffCompression enum).
        tiffOptions.TiffCompression = TiffCompression.Lzw;
        // Optional: define the resolution (dots per inch) for the output image.
        tiffOptions.Resolution = 300;

        // Save the document as a multi‑page TIFF file.
        doc.Save("Output.tiff", tiffOptions);
    }
}
