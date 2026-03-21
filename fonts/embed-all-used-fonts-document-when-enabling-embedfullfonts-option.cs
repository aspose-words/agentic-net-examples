using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedAllFontsExample
{
    static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a test document with embedded fonts.");

        // Configure PDF save options to embed full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF with all used fonts fully embedded.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}
