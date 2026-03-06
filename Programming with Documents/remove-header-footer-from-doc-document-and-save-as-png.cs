using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Configure image save options to render the document as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: set a higher resolution for better quality (optional).
        pngOptions.Resolution = 300;

        // Save the document as a PNG image. By default only the first page is rendered.
        doc.Save("output.png", pngOptions);
    }
}
