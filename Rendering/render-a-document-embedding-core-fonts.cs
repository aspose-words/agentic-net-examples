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

        // Use core TrueType fonts that will be substituted with PDF Type 1 fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Courier New";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Configure PDF save options to replace the fonts with core PDF fonts.
        PdfSaveOptions options = new PdfSaveOptions();
        options.UseCoreFonts = true; // Enable core font substitution.

        // Save the document as PDF using the configured options.
        doc.Save("PdfSaveOptions.CoreFonts.pdf", options);
    }
}
