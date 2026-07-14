using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class InstallTrueTypeFontsLinux
{
    public static void Main()
    {
        // Define directories for artifacts, source fonts, and installed fonts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string sourceFontsDir = Path.Combine(artifactsDir, "SourceFonts");
        string installedFontsDir = Path.Combine(artifactsDir, "InstalledFonts");

        // Ensure directories exist.
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(sourceFontsDir);
        Directory.CreateDirectory(installedFontsDir);

        // Create a dummy TrueType font file in the source directory.
        // In a real scenario you would copy actual .ttf files from another location.
        string dummyFontFileName = "DummyFont.ttf";
        string dummyFontPath = Path.Combine(sourceFontsDir, dummyFontFileName);
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal TTF header (just to have a non‑empty file).
            byte[] minimalTtfHeader = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x80 };
            File.WriteAllBytes(dummyFontPath, minimalTtfHeader);
        }

        // Copy all .ttf files from the source folder to the installed fonts folder.
        foreach (string ttfFile in Directory.GetFiles(sourceFontsDir, "*.ttf"))
        {
            string destFile = Path.Combine(installedFontsDir, Path.GetFileName(ttfFile));
            File.Copy(ttfFile, destFile, true);
        }

        // Point Aspose.Words to the folder that now contains the TrueType fonts.
        FontSettings.DefaultInstance.SetFontsFolder(installedFontsDir, recursive: true);

        // Create a sample document that uses the dummy font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text should be rendered with the DummyFont.");

        // Save the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "Output.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure fonts are embedded (subset) so we can verify embedding.
            EmbedFullFonts = false
        };
        doc.Save(pdfPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF output.", pdfPath);

        // Simple validation: check the PDF content for an embedded font marker (e.g., "/FontFile").
        string pdfText = File.ReadAllText(pdfPath);
        bool containsFontMarker = pdfText.Contains("/FontFile") || pdfText.Contains("/FontFile2") || pdfText.Contains("/FontFile3");
        if (!containsFontMarker)
            throw new InvalidOperationException("The generated PDF does not contain embedded font markers.");

        // Cleanup: (optional) reset font sources to original state.
        FontSettings.DefaultInstance.ResetFontSources();
    }
}
