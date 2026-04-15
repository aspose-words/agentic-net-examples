using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that contains OpenType ligatures (e.g., Arial).
        builder.Font.Name = "Arial";

        // Add text that includes ligature candidates.
        builder.Writeln("Office");          // contains "ff"
        builder.Writeln("affinity");        // contains "ff" and "fi"
        builder.Writeln("efficient");       // contains "ff" and "fi"
        builder.Writeln("fluff");           // contains "fl"

        // Configure font settings to ensure the font can be found.
        FontSettings fontSettings = new FontSettings();
        string systemFonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(systemFonts, true);
        doc.FontSettings = fontSettings;

        // Set PDF save options to embed the full font, preserving OpenType features.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "OpenTypeLigatures.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // Optional: reload the PDF to ensure it is readable.
        Document loadedPdf = new Document(pdfPath);
        // No further actions required; the example completes here.
    }
}
