using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string currentDir = Directory.GetCurrentDirectory();
        string fontsDir = Path.Combine(currentDir, "Fonts");
        Directory.CreateDirectory(fontsDir);
        string outputPdf = Path.Combine(currentDir, "DocumentWithCustomFont.pdf");

        // Locate a system TrueType font (Arial) and copy it to the local Fonts folder.
        // This ensures the example is self‑contained and does not rely on an external path.
        string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string sourceFontPath = Path.Combine(systemFontsPath, "arial.ttf");
        if (!File.Exists(sourceFontPath))
            throw new FileNotFoundException("System font not found: " + sourceFontPath);

        string localFontPath = Path.Combine(fontsDir, "arial.ttf");
        if (!File.Exists(localFontPath))
            File.Copy(sourceFontPath, localFontPath, overwrite: true);

        // Load the font bytes into memory.
        byte[] fontData = File.ReadAllBytes(localFontPath);
        MemoryFontSource memoryFont = new MemoryFontSource(fontData);

        // Configure FontSettings to use the in‑memory font source.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsSources(new FontSourceBase[] { memoryFont });

        // Create a blank document and assign the custom FontSettings.
        Document doc = new Document();
        doc.FontSettings = fontSettings;

        // Add some text using the custom font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Font name matches the loaded font.
        builder.Writeln("This paragraph is rendered with a custom embedded font.");

        // Set PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            EmbedFullFonts = true // Embed the full font file (no subsetting).
        };

        // Save the document as PDF.
        doc.Save(outputPdf, pdfOptions);

        // Validate that the PDF was created and is not empty.
        if (!File.Exists(outputPdf) || new FileInfo(outputPdf).Length == 0)
            throw new InvalidOperationException("PDF conversion failed: output file is missing or empty.");

        // Inform the user (optional, not required for non‑interactive execution).
        Console.WriteLine("PDF successfully created at: " + outputPdf);
    }
}
