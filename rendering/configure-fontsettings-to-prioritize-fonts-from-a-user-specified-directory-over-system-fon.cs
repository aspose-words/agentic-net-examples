using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);
        string customFontsDir = Path.Combine(artifactsDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Determine a system font to copy into the custom folder (if possible).
        string systemFontFolder = SystemFontSource.GetSystemFontFolders().FirstOrDefault();
        string fontFilePath = null;
        string fontName = "Arial"; // fallback font name.

        if (!string.IsNullOrEmpty(systemFontFolder))
        {
            // Pick the first TrueType font file we can find.
            string[] ttfFiles = Directory.GetFiles(systemFontFolder, "*.ttf");
            if (ttfFiles.Length > 0)
            {
                fontFilePath = ttfFiles[0];
                fontName = Path.GetFileNameWithoutExtension(fontFilePath);
                // Copy the font file into the custom folder.
                string destPath = Path.Combine(customFontsDir, Path.GetFileName(fontFilePath));
                File.Copy(fontFilePath, destPath, true);
            }
        }

        // Create a simple document that uses the selected font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = fontName;
        builder.Writeln($"This text is rendered using the font \"{fontName}\".");

        // Configure FontSettings to prioritize the custom folder over system fonts.
        FontSettings fontSettings = new FontSettings();

        // Create a folder font source for the custom fonts directory.
        FolderFontSource folderSource = new FolderFontSource(customFontsDir, false);
        // Keep the default system font source.
        SystemFontSource systemSource = new SystemFontSource();

        // Set the sources with the custom folder first (higher priority).
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource, systemSource });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Optionally, you could inspect the PDF file size or content here.
        Console.WriteLine($"PDF successfully saved to: {pdfPath}");
    }
}
