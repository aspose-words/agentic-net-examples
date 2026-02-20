using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class EmbedMissingFontsAndSavePdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure font settings to search for missing fonts on the system.
        // The second parameter (true) tells Aspose.Words to also search subfolders.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(@"C:\Windows\Fonts", true);
        doc.FontSettings = fontSettings;

        // Set PDF save options to embed all fonts (including those that were missing).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed every font used in the document.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,

            // Embed the full font data (not just subsets) for maximum fidelity.
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
