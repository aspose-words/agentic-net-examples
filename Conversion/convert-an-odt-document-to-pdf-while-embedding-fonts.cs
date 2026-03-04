using System;
using Aspose.Words;
using Aspose.Words.Saving;

class OdtToPdfConverter
{
    static void Main()
    {
        // Path to the source ODT document.
        string sourcePath = "input.odt";

        // Path where the resulting PDF will be saved.
        string destinationPath = "output.pdf";

        // Load the ODT document.
        Document document = new Document(sourcePath);

        // Configure PDF save options to embed fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the full font files (no subsetting) so that all glyphs are available.
            EmbedFullFonts = true,

            // Ensure all fonts, including standard ones, are embedded.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,

            // Allow embedding of PostScript outlines when embedding TrueType fonts.
            AllowEmbeddingPostScriptFonts = true
        };

        // Save the document as PDF using the configured options.
        document.Save(destinationPath, pdfOptions);
    }
}
