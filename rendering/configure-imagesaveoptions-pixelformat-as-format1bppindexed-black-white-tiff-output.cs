using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Configure the pixel format to 1‑bit indexed (black‑and‑white).
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use CCITT4 compression which is suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a black‑and‑white TIFF file.
        doc.Save("Output.tiff", options);
    }
}
