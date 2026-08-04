using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a folder that simulates a network share.
        string networkFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "NetworkFonts");
        Directory.CreateDirectory(networkFontsFolder);

        // Locate a TrueType font on the local system and copy it to the simulated network folder.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length == 0)
            throw new InvalidOperationException("No system font folders found.");

        string sourceFontPath = null;
        foreach (var folder in systemFontFolders)
        {
            var ttfFiles = Directory.GetFiles(folder, "*.ttf");
            if (ttfFiles.Length > 0)
            {
                sourceFontPath = ttfFiles[0];
                break;
            }
        }

        if (sourceFontPath == null)
            throw new FileNotFoundException("No TrueType font file found on the system.");

        string destFontPath = Path.Combine(networkFontsFolder, Path.GetFileName(sourceFontPath));
        File.Copy(sourceFontPath, destFontPath, true);

        // Create custom FontSettings pointing to the network folder.
        FontSettings customFontSettings = new FontSettings();
        customFontSettings.SetFontsFolder(networkFontsFolder, true);

        // Build a simple document that uses a font available in the copied folder.
        Document doc = new Document();
        doc.FontSettings = customFontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = Path.GetFileNameWithoutExtension(destFontPath);
        builder.Writeln("This text is rendered using a font loaded from the network folder.");

        // Save the document to PDF to trigger font resolution.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.pdf");
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary files (optional).
        // File.Delete(outputPath);
        // Directory.Delete(networkFontsFolder, true);
    }
}
