using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text using standard Windows fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("This text uses Arial, a standard Windows font.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text uses Times New Roman, another standard Windows font.");

        // Add text using a non‑standard font.
        builder.Font.Name = "Courier New";
        builder.Writeln("This text uses Courier New, a non‑standard font for this example.");

        // Configure PDF save options.
        PdfSaveOptions options = new PdfSaveOptions();

        // Substitute standard fonts with core PDF Type 1 fonts.
        options.UseCoreFonts = true;

        // Embed only non‑standard fonts (skip embedding Arial and Times New Roman).
        options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNonstandard;

        // Embed full glyphs for the non‑standard font (optional).
        options.EmbedFullFonts = true;

        // Save the document as PDF with the specified options.
        doc.Save("Output.pdf", options);
    }
}
