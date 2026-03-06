using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder where the output PDF will be saved.
        string artifactsDir = @"C:\Artifacts\";

        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text with different fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Standard Windows font – will be substituted by a core PDF font.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");

        // Non‑standard font – will be embedded according to the embedding mode.
        builder.Font.Name = "Courier New";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Configure PDF save options.
        PdfSaveOptions options = new PdfSaveOptions();

        // Substitute TrueType fonts Arial, Times New Roman, Courier New and Symbol
        // with the corresponding core PDF Type 1 fonts.
        options.UseCoreFonts = true;

        // Embed only non‑standard fonts (skip embedding standard Windows fonts).
        options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNonstandard;

        // Save the document as PDF using the configured options.
        doc.Save(artifactsDir + "RenderCoreAndWindowsFonts.pdf", options);
    }
}
