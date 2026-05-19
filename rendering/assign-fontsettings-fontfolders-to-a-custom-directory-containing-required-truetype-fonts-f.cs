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
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string customFontsDir = Path.Combine(artifactsDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Locate a TrueType font from the system and copy it to the custom folder.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length == 0)
            throw new Exception("No system font folders were found.");

        string sourceFontPath = Directory.GetFiles(systemFontFolders[0], "*.ttf", SearchOption.AllDirectories).FirstOrDefault();
        if (sourceFontPath != null)
        {
            string destFontPath = Path.Combine(customFontsDir, Path.GetFileName(sourceFontPath));
            File.Copy(sourceFontPath, destFontPath, true);
        }

        // Create FontSettings that point to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontsDir, true);

        // Create a simple document that uses a font from the custom folder.
        Document doc = new Document();
        doc.FontSettings = fontSettings;
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Determine a font family name that is available in the custom folder.
        FolderFontSource folderSource = new FolderFontSource(customFontsDir, true);
        string fontFamily = folderSource.GetAvailableFonts().FirstOrDefault()?.FontFamilyName ?? "Arial";

        builder.Font.Name = fontFamily;
        builder.Writeln("This text is rendered using the custom font folder.");

        // Save the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "CustomFontDocument.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("Failed to create the PDF output.");

        // Optionally, clean up (commented out to keep output files for inspection).
        // Directory.Delete(customFontsDir, true);
    }
}
