using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Create save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the horizontal and vertical DPI to 300.
        options.HorizontalResolution = 300f;
        options.VerticalResolution = 300f;
        // The Resolution property could also be used to set both at once:
        // options.Resolution = 300f;

        // Save the document as a TIFF image with the specified DPI.
        doc.Save("output.tiff", options);
    }
}
