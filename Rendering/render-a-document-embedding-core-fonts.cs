using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentWithCoreFonts
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text with fonts that can be substituted.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Courier New";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Configure PDF save options to substitute TrueType fonts with core PDF Type 1 fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseCoreFonts = true // Enable core font substitution.
        };

        // Save the document as PDF using the configured options.
        doc.Save("RenderWithCoreFonts.pdf", pdfOptions);
    }
}
