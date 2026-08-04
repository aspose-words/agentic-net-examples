using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Configure font settings to use Arial Unicode MS as the default fallback font.
        FontSettings fontSettings = new FontSettings();
        DefaultFontSubstitutionRule defaultSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
        defaultSubstitution.DefaultFontName = "Arial Unicode MS"; // Fallback font.
        // The rule is enabled by default, but we set it explicitly for clarity.
        defaultSubstitution.Enabled = true;

        // Assign the font settings to the document.
        doc.FontSettings = fontSettings;

        // Build document content using a font that does not exist on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font";
        builder.Writeln("This line uses a missing font. Characters: 漢字, 😊, 𝔘𝔫𝔦𝔠𝔬𝔡𝔢.");

        // Save the document to PDF. The missing font will be substituted with Arial Unicode MS.
        string outputPath = Path.Combine(artifactsDir, "FontFallbackExample.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document saved successfully to: " + outputPath);
    }
}
