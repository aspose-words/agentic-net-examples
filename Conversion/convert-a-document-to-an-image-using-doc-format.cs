using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Configure image save options – render to PNG.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: render only the first page (pages are zero‑based).
        imageOptions.PageSet = new PageSet(0);

        // Save the rendered page as an image file.
        doc.Save("Output.png", imageOptions);
    }
}
