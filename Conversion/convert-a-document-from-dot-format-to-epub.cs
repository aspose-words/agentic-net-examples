using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document in DOT format.
        Document doc = new Document("input.dot");

        // Configure save options to produce an EPUB file.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Save the document as EPUB using the configured options.
        doc.Save("output.epub", epubOptions);
    }
}
