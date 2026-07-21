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
        // Prepare directories for output and custom fonts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string customFontsDir = Path.Combine(artifactsDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Attempt to copy a TrueType font from a system font folder into the custom folder.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length > 0)
        {
            string sysFolder = systemFontFolders[0];
            string[] ttfFiles = Directory.GetFiles(sysFolder, "*.ttf");
            if (ttfFiles.Length > 0)
            {
                string sourceFontPath = ttfFiles[0];
                string destFontPath = Path.Combine(customFontsDir, Path.GetFileName(sourceFontPath));
                File.Copy(sourceFontPath, destFontPath, true);
            }
        }

        // Determine a font name that exists in the custom folder (fallback to Arial if none found).
        string fontFilePath = Directory.GetFiles(customFontsDir, "*.ttf").FirstOrDefault();
        string fontName = fontFilePath != null
            ? Path.GetFileNameWithoutExtension(fontFilePath)
            : "Arial";

        // Create a simple document that uses the selected font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = fontName;
        builder.Writeln($"This text uses the font \"{fontName}\" loaded from a custom folder.");

        // Assign FontSettings to point to the custom fonts directory.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontsDir, false);
        doc.FontSettings = fontSettings;

        // Render the document to PDF, embedding fonts.
        string pdfPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optional: indicate success.
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}
