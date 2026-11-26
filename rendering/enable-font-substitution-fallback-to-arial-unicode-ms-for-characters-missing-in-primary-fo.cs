using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a primary font that does not contain Chinese characters.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("Hello world!");
        builder.Writeln("中文字符测试"); // Chinese characters.

        // Configure font settings with fallback to Arial Unicode MS.
        FontSettings fontSettings = new FontSettings();

        // Ensure Aspose.Words can locate system fonts.
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, true);

        // Set fallback for missing glyphs.
        var defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.DefaultFontName = "Arial Unicode MS";

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = "RenderedWithFallback.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF file was not created.", outputPath);

        Console.WriteLine($"Document rendered successfully to '{outputPath}'.");
    }
}
