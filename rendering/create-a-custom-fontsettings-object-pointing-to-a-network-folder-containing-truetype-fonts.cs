using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Simulate a network folder that contains TrueType fonts.
        string networkFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "NetworkFonts");
        Directory.CreateDirectory(networkFontsFolder);

        // Copy a known system font (e.g., Arial) into the simulated network folder.
        // This ensures the folder actually contains a usable TrueType font.
        string systemFontPath = FindSystemFont("Arial.ttf");
        if (systemFontPath != null)
        {
            string destination = Path.Combine(networkFontsFolder, Path.GetFileName(systemFontPath));
            File.Copy(systemFontPath, destination, true);
        }

        // Create a simple document that uses the copied font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered with a font loaded from a network folder.");

        // Configure FontSettings to point to the network font folder.
        FontSettings fontSettings = new FontSettings();
        // 'true' enables recursive scanning of subfolders.
        fontSettings.SetFontsFolder(networkFontsFolder, true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF to trigger layout and font resolution.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output PDF.");
    }

    // Helper method that searches system font folders for a specific font file.
    private static string FindSystemFont(string fontFileName)
    {
        foreach (string folder in SystemFontSource.GetSystemFontFolders())
        {
            string candidate = Path.Combine(folder, fontFileName);
            if (File.Exists(candidate))
                return candidate;
        }
        return null;
    }
}
