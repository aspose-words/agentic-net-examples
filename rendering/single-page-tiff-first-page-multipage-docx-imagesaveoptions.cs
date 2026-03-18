using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX which contains multiple pages.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Render only the first page (zero‑based index) of the document.
        options.PageSet = new PageSet(0);

        // Save the first page as a single‑page TIFF file.
        doc.Save("FirstPage.tiff", options);
    }
}
