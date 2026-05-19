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
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string customFontsDir = Path.Combine(outputDir, "MyFonts");
        Directory.CreateDirectory(customFontsDir);

        // Attempt to copy a TrueType font from a system font folder into the custom folder.
        // This ensures the custom folder contains a usable font for the example.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length > 0)
        {
            string systemFolder = systemFontFolders[0];
            string sourceFontPath = Directory.GetFiles(systemFolder, "*.ttf").FirstOrDefault();
            if (!string.IsNullOrEmpty(sourceFontPath))
            {
                string destFontPath = Path.Combine(customFontsDir, Path.GetFileName(sourceFontPath));
                File.Copy(sourceFontPath, destFontPath, true);
            }
        }

        // Create a simple document that uses a font which should be resolved from the custom folder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Common font; we copied its file if possible.
        builder.Writeln("This text should be rendered using the font from the custom directory.");

        // Configure FontSettings to prioritize the custom font folder.
        FontSettings fontSettings = new FontSettings();
        // The folder will be scanned recursively; it becomes the sole font source,
        // thus taking precedence over system fonts.
        fontSettings.SetFontsFolder(customFontsDir, true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(outputDir, "Result.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output.");

        Console.WriteLine($"PDF rendered successfully to: {pdfPath}");
    }
}
