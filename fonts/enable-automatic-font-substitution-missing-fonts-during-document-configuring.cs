using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontSubstitutionExample
{
    static void Main()
    {
        // Output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.pdf");

        // Configure font substitution settings.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Create a new document and apply the font settings.
        Document doc = new Document();
        doc.FontSettings = fontSettings;

        // Add some text that uses a font that is likely missing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and will be substituted with Arial.");

        // Save the document; missing fonts will be substituted automatically.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
