using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document from the file system.
        Document doc = new Document("Input.doc");

        // Configure image save options – render pages as PNG images.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Example: set the resolution to 300 DPI for higher quality.
        options.Resolution = 300;

        // Save the document (first page) as an image file.
        doc.Save("Output.png", options);
    }
}
