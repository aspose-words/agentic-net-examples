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

        // Copy a TrueType font from the system fonts folder into the simulated network folder.
        // This ensures the folder actually contains a .ttf file for the example.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length > 0)
        {
            string[] ttfFiles = Directory.GetFiles(systemFontFolders[0], "*.ttf");
            if (ttfFiles.Length > 0)
            {
                string sourceFont = ttfFiles[0];
                string destFont = Path.Combine(networkFontsFolder, Path.GetFileName(sourceFont));
                File.Copy(sourceFont, destFont, true);
            }
        }

        // Create custom FontSettings that point to the network folder.
        FontSettings fontSettings = new FontSettings();
        // The second argument 'true' enables recursive scanning of subfolders.
        fontSettings.SetFontsFolder(networkFontsFolder, true);

        // Build a simple document that uses a font expected to be found in the network folder.
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered using the Arial font loaded from the network folder.");

        // Save the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.pdf");
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");
    }
}
