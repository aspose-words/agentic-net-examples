using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Create save options for EPUB output.
        // The constructor that takes a SaveFormat allows us to specify the target format directly.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Save the document as a single EPUB file.
        doc.Save("output.epub", epubOptions);
    }
}
