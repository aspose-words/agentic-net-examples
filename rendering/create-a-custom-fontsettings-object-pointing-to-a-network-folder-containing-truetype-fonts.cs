using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Simulate a network folder that contains TrueType fonts.
        // In a real scenario this would be a UNC path like \\server\share\fonts.
        string networkFontsDir = Path.Combine(artifactsDir, "NetworkFonts");
        Directory.CreateDirectory(networkFontsDir);

        // Optionally copy a system TrueType font into the simulated network folder
        // so that the document can actually use a custom font.
        string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string sampleFontPath = Directory.GetFiles(systemFontsPath, "*.ttf")
                                         .FirstOrDefault();

        string copiedFontPath = null;
        if (sampleFontPath != null)
        {
            string destPath = Path.Combine(networkFontsDir, Path.GetFileName(sampleFontPath));
            File.Copy(sampleFontPath, destPath, true);
            copiedFontPath = destPath;
        }

        // Create custom FontSettings and point it to the network folder.
        FontSettings customFontSettings = new FontSettings();
        // The second argument indicates whether to scan subfolders recursively.
        customFontSettings.SetFontsFolder(networkFontsDir, recursive: true);

        // Build a simple document that uses a font from the network folder (if we copied one).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a known system font.
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses the system font Arial.");

        // If we have a copied font, use it to demonstrate the custom font source.
        if (copiedFontPath != null)
        {
            string fontName = Path.GetFileNameWithoutExtension(copiedFontPath);
            builder.Font.Name = fontName;
            builder.Writeln($"This line uses the custom font \"{fontName}\" loaded from the network folder.");
        }

        // Assign the custom FontSettings to the document.
        doc.FontSettings = customFontSettings;

        // Save the document to PDF to force layout and font resolution.
        string pdfPath = Path.Combine(artifactsDir, "CustomFontSettings.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // The example finishes without requiring any user interaction.
    }
}
