using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "FallbackExample.pdf");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using a normal font.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses a standard font.");

        // Write text using a font that is unlikely to be available and contains characters
        // that are not present in the primary font (e.g., Chinese characters).
        builder.Font.Name = "Missing Font";
        builder.Writeln("Missing Font line with Unicode characters: 漢字");

        // Configure font settings to use Arial Unicode MS as the default substitution font.
        FontSettings fontSettings = new FontSettings();
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.DefaultFontName = "Arial Unicode MS";
        defaultSubstitution.Enabled = true;

        // Assign the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new Exception("Failed to create the PDF output.");

        // Optionally, you could add further validation here (e.g., checking file size).
    }
}
