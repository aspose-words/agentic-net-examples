using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define the folder where the output PDF will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Set up PDF save options to embed the full font files (no subsetting).
        PdfSaveOptions options = new PdfSaveOptions();
        options.EmbedFullFonts = true;

        // Save the document as a PDF with full font embedding.
        string outputPath = Path.Combine(artifactsDir, "FullFontsEmbedded.pdf");
        doc.Save(outputPath, options);
    }
}
