using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define source and target font directories.
        string sourceFontDir = Path.Combine(Directory.GetCurrentDirectory(), "sourceFonts");
        string targetFontDir = Path.Combine(Directory.GetCurrentDirectory(), "fonts");

        // Ensure directories exist.
        Directory.CreateDirectory(sourceFontDir);
        Directory.CreateDirectory(targetFontDir);

        // Create a dummy TrueType font file in the source directory.
        string dummyFontFileName = "DummyFont.ttf";
        string sourceFontPath = Path.Combine(sourceFontDir, dummyFontFileName);
        if (!File.Exists(sourceFontPath))
        {
            // Write minimal placeholder bytes (not a real font, but sufficient for copy demonstration).
            byte[] placeholder = new byte[] { 0x00, 0x01, 0x00, 0x00 };
            File.WriteAllBytes(sourceFontPath, placeholder);
        }

        // Copy all .ttf files from source to target directory.
        foreach (string ttfPath in Directory.GetFiles(sourceFontDir, "*.ttf"))
        {
            string destPath = Path.Combine(targetFontDir, Path.GetFileName(ttfPath));
            File.Copy(ttfPath, destPath, true);
        }

        // Verify that the font was copied.
        string installedFontPath = Path.Combine(targetFontDir, dummyFontFileName);
        if (!File.Exists(installedFontPath))
            throw new Exception("Failed to install the TrueType font.");

        // Configure FontSettings to use the installed fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(targetFontDir, false);

        // Create a simple document and assign the FontSettings.
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Hello, world! This document uses installed fonts."));

        // Render the document to PDF.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("PDF rendering failed.");

        // Optionally, output confirmation (no interactive wait).
        Console.WriteLine("Font installation and PDF rendering completed successfully.");
    }
}
