using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedTrueTypeFontsToPdf
{
    static void Main()
    {
        // Path to the input Word document.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Input.docx");

        // Path to the folder where the output PDF will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "DocumentWithEmbeddedFonts.pdf");

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create PDF save options and enable full font embedding.
        PdfSaveOptions options = new PdfSaveOptions
        {
            // When true, every glyph of every TrueType font used in the document is embedded.
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, options);
    }
}
