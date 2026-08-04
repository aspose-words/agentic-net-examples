using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample output and the custom font folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string customFontDir = Path.Combine(outputDir, "UserFonts");
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(customFontDir);

        // Locate a system font file to copy into the custom folder.
        // This ensures we have a valid TrueType font for the demonstration.
        string systemFontFile = null;
        foreach (string folder in SystemFontSource.GetSystemFontFolders())
        {
            systemFontFile = Directory.GetFiles(folder, "*.ttf").FirstOrDefault();
            if (systemFontFile != null)
                break;
        }

        if (systemFontFile == null)
            throw new FileNotFoundException("No TrueType font file found in system font folders.");

        // Copy the font file to the custom font directory.
        string copiedFontPath = Path.Combine(customFontDir, Path.GetFileName(systemFontFile));
        File.Copy(systemFontFile, copiedFontPath, true);

        // Determine the font name (without extension) to use in the document.
        string fontName = Path.GetFileNameWithoutExtension(copiedFontPath);

        // Create a simple document that uses the selected font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = fontName;
        builder.Writeln($"This text is rendered using the font \"{fontName}\" from the custom folder.");

        // Configure FontSettings to prioritize the custom font folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(customFontDir, recursive: true);
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");
        doc.Save(pdfPath);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, output the path of the generated file.
        Console.WriteLine($"Document rendered successfully: {pdfPath}");
    }
}
