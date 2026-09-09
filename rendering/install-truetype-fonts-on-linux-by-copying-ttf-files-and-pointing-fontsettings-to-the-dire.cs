using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class InstallTrueTypeFontsOnLinux
{
    public static void Main()
    {
        // Define folders for the example.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string sourceFontsDir = Path.Combine(baseDir, "SourceFonts");
        string linuxFontsDir = Path.Combine(baseDir, "LinuxFonts");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(sourceFontsDir);
        Directory.CreateDirectory(linuxFontsDir);

        // Create a dummy TrueType font file in the source folder.
        // In a real scenario this would be an actual .ttf file.
        string dummyFontPath = Path.Combine(sourceFontsDir, "DummyFont.ttf");
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal TTF header (just to have a non‑empty file).
            byte[] dummyTtfHeader = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x80 };
            File.WriteAllBytes(dummyFontPath, dummyTtfHeader);
        }

        // Copy the font file to the Linux fonts directory (simulating installation).
        string installedFontPath = Path.Combine(linuxFontsDir, "DummyFont.ttf");
        File.Copy(dummyFontPath, installedFontPath, true);

        // Point Aspose.Words to the directory that contains the installed fonts.
        // The second argument 'true' enables recursive scanning of subfolders.
        FontSettings.DefaultInstance.SetFontsFolder(linuxFontsDir, true);

        // Build a simple document that uses the dummy font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text is rendered with the installed TrueType font.");

        // Save the document to PDF.
        string pdfPath = Path.Combine(baseDir, "RenderedDocument.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed full fonts to ensure the font is included in the PDF.
            EmbedFullFonts = true
        };
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF output was not generated.", pdfPath);

        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
