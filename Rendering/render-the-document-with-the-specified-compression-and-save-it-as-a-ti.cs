using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create an ImageSaveOptions instance for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Specify the compression type to use when saving the TIFF.
        // Available values: TiffCompression.None, Rle, Lzw, Ccitt3, Ccitt4.
        options.TiffCompression = TiffCompression.Lzw;

        // Save the rendered document as a TIFF file using the configured options.
        doc.Save("Output.tiff", options);
    }
}
