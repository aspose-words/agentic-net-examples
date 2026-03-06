using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using TrueType fonts that can be substituted.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world! This line uses Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("Hello world! This line uses Times New Roman.");

        // Configure PDF save options to replace the TrueType fonts with core PDF Type 1 fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.UseCoreFonts = true; // Enable font substitution.

        // Save the document as a PDF using the configured options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
