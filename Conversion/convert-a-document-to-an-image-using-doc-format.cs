using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document from the file system.
        Document doc = new Document("input.doc");

        // Configure image save options – render pages as PNG images.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Example: set a higher resolution for better quality.
        options.Resolution = 300; // DPI

        // Save the document (first page by default) as an image file.
        doc.Save("output.png", options);
    }
}
