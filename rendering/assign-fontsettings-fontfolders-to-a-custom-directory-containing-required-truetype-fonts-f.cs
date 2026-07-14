using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define a custom folder that will hold the TrueType fonts.
        string customFontsDir = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Try to locate a TrueType font file from the system font folders.
        string systemFontPath = null;
        foreach (string folder in SystemFontSource.GetSystemFontFolders())
        {
            if (Directory.Exists(folder))
            {
                string[] ttfFiles = Directory.GetFiles(folder, "*.ttf");
                if (ttfFiles.Length > 0)
                {
                    systemFontPath = ttfFiles[0];
                    break;
                }
            }
        }

        // If a system font was found, copy it into the custom folder.
        if (systemFontPath != null)
        {
            string destPath = Path.Combine(customFontsDir, Path.GetFileName(systemFontPath));
            File.Copy(systemFontPath, destPath, true);
        }
        else
        {
            // No system font found – the example will still demonstrate assigning the folder.
            Console.WriteLine("No system TrueType font found; proceeding with an empty custom font folder.");
        }

        // Determine the name of the font we just placed (if any).
        string fontName = "Arial"; // Fallback font name.
        try
        {
            var folderSource = new FolderFontSource(customFontsDir, true);
            var availableFonts = folderSource.GetAvailableFonts();
            if (availableFonts.Any())
                fontName = availableFonts.First().FullFontName;
        }
        catch
        {
            // If any error occurs while reading the custom folder, keep the fallback font.
        }

        // Create a new document and configure its FontSettings to use the custom folder.
        Document doc = new Document();
        doc.FontSettings = new FontSettings();
        doc.FontSettings.SetFontsFolder(customFontsDir, true);

        // Write some text using the font from the custom folder (or fallback).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = fontName;
        builder.Writeln($"This text is rendered with the font \"{fontName}\" from the custom folder.");

        // Render the document to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedOutput.pdf");
        doc.Save(outputPath);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The rendered PDF file was not created.");

        // Clean up (optional): delete the custom font folder and its contents.
        // Directory.Delete(customFontsDir, true);
    }
}
