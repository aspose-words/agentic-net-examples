using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Configure image save options to render the document as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: set a higher resolution for better quality (optional).
        pngOptions.Resolution = 300;

        // Save the document. When saving to an image format only the first page is rendered.
        doc.Save("Output.png", pngOptions);
    }
}
