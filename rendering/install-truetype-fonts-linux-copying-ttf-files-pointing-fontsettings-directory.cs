using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class InstallFontsOnLinux
{
    static void Main()
    {
        // Directory containing .ttf files to install (adjust if you have custom fonts).
        string sourceFontsDir = Path.Combine(Environment.CurrentDirectory, "fonts");
        // Directory where Aspose.Words will look for fonts.
        string targetFontsDir = Path.Combine(Environment.CurrentDirectory, "app_fonts");

        // Ensure the target directory exists.
        Directory.CreateDirectory(targetFontsDir);

        // Copy all TrueType font files (*.ttf) from the source to the target directory, if the source exists.
        if (Directory.Exists(sourceFontsDir))
        {
            foreach (string fontFilePath in Directory.GetFiles(sourceFontsDir, "*.ttf", SearchOption.AllDirectories))
            {
                string destPath = Path.Combine(targetFontsDir, Path.GetFileName(fontFilePath));
                File.Copy(fontFilePath, destPath, true);
            }
        }
        else
        {
            Console.WriteLine($"Source fonts directory not found: {sourceFontsDir}. Continuing without custom fonts.");
        }

        // Configure Aspose.Words to use the target fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(targetFontsDir, true);
        // Optional: set a fallback font name if needed (commented out because the property may not exist in some versions).
        // fontSettings.SubstitutionSettings.DefaultFontName = "Arial";

        // Create or load a document.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "input.docx");
        Document doc;
        if (File.Exists(inputPath))
        {
            LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };
            doc = new Document(inputPath, loadOptions);
        }
        else
        {
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This document uses the configured font settings.");
        }

        // Ensure the document uses the configured FontSettings.
        doc.FontSettings = fontSettings;

        // Enable font substitution and embedding options.
        doc.FontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;
        doc.FontInfos.EmbedTrueTypeFonts = true;
        doc.FontInfos.EmbedSystemFonts = true;
        doc.FontInfos.SaveSubsetFonts = true;

        // Save the processed document as PDF.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.pdf");
        try
        {
            doc.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save document: {ex.Message}");
        }
    }
}
