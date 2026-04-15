using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the resulting PDF.
        string pdfPath = Path.Combine(outputDir, "SampleSubset.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document in memory.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a TrueType font that is typically available on the system.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello World! This text will be rendered to PDF.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // -----------------------------------------------------------------
        // 2. Ensure Aspose.Words can locate the fonts.
        //    Save the original font sources so we can restore them later.
        // -----------------------------------------------------------------
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add the OS fonts folder as a font source (recursive search).
        string systemFontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        FontSettings.DefaultInstance.SetFontsFolder(systemFontsFolder, true);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options to embed only the glyphs used
        //    (font subsetting). The default value of EmbedFullFonts is false,
        //    but we set it explicitly for clarity.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = false   // Subset fonts – only used glyphs are embedded.
        };

        // -----------------------------------------------------------------
        // 4. Save the document as PDF.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 6. Restore the original font sources to avoid side‑effects.
        // -----------------------------------------------------------------
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
