using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a user‑specified font folder.
        string userFontsDir = Path.Combine(artifactsDir, "UserFonts");
        Directory.CreateDirectory(userFontsDir);

        // Locate a TrueType font from the system font folders.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length == 0)
            throw new InvalidOperationException("No system font folders were found.");

        // Find the first .ttf file in the first system font folder.
        string systemFontPath = Directory.GetFiles(systemFontFolders[0], "*.ttf", SearchOption.TopDirectoryOnly).FirstOrDefault();
        if (systemFontPath == null)
            throw new FileNotFoundException("No TrueType font file found in the system font folder.");

        // Copy the font file into the user‑specified folder.
        string copiedFontPath = Path.Combine(userFontsDir, Path.GetFileName(systemFontPath));
        File.Copy(systemFontPath, copiedFontPath, true);

        // Build a simple document that uses the copied font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use the font name extracted from the file name (e.g., "Arial").
        string fontName = Path.GetFileNameWithoutExtension(copiedFontPath);
        builder.Font.Name = fontName;
        builder.Writeln($"This text is rendered with the font \"{fontName}\" from the user‑specified folder.");

        // Configure FontSettings to prioritize the user‑specified folder.
        FontSettings fontSettings = new FontSettings();
        // This call replaces the default font sources with the folder we provide.
        fontSettings.SetFontsFolder(userFontsDir, true);
        // Assign the settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The rendered PDF was not created.");

        // The example finishes without requiring user interaction.
    }
}
