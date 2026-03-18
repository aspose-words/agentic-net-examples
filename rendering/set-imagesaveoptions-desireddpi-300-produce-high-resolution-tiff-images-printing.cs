using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the resolution (DPI) to 300 for high‑resolution printing.
        // The correct property is Resolution, not DesiredDpi.
        options.Resolution = 300;

        // Save the document as a TIFF image using the specified DPI.
        doc.Save("output.tiff", options);
    }
}
