using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class InstallTrueTypeFontsExample
{
    public static void Main()
    {
        // Define source and target font directories relative to the current working directory.
        string currentDir = Directory.GetCurrentDirectory();
        string sourceFontsDir = Path.Combine(currentDir, "SourceFonts");
        string installedFontsDir = Path.Combine(currentDir, "InstalledFonts");

        // Ensure the directories exist.
        Directory.CreateDirectory(sourceFontsDir);
        Directory.CreateDirectory(installedFontsDir);

        // Create a dummy TrueType font file in the source directory if none exists.
        // In a real scenario, you would copy actual .ttf files.
        string dummyFontPath = Path.Combine(sourceFontsDir, "DummyFont.ttf");
        if (!File.Exists(dummyFontPath))
        {
            // Write a minimal placeholder byte array (not a valid font, but sufficient for the example).
            File.WriteAllBytes(dummyFontPath, new byte[] { 0x00, 0x01, 0x00, 0x00 });
        }

        // Copy all .ttf files from the source directory to the installed fonts directory.
        foreach (string fontFile in Directory.GetFiles(sourceFontsDir, "*.ttf"))
        {
            string destFile = Path.Combine(installedFontsDir, Path.GetFileName(fontFile));
            File.Copy(fontFile, destFile, true);
        }

        // Configure Aspose.Words to use the installed fonts directory.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(installedFontsDir, recursive: true);

        // Create a simple document that uses the dummy font.
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "DummyFont";
        builder.Writeln("This text is rendered with the DummyFont.");

        // Save the document to PDF.
        string outputPdfPath = Path.Combine(currentDir, "Output.pdf");
        doc.Save(outputPdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("Failed to create the output PDF file.");

        Console.WriteLine($"PDF successfully saved to: {outputPdfPath}");
    }
}
