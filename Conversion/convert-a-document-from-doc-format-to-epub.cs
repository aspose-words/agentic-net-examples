using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file.
        // Aspose.Words automatically detects the format, so no explicit LoadFormat is required.
        Document doc = new Document("input.doc");

        // Create save options for EPUB output.
        // HtmlSaveOptions can be used for HTML, MHTML, EPUB, AZW3 and MOBI formats.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Optional: configure additional EPUB-specific options here, e.g. navigation map depth.
        // epubOptions.NavigationMapLevel = 3;

        // Save the document as EPUB using the specified options.
        doc.Save("output.epub", epubOptions);
    }
}
