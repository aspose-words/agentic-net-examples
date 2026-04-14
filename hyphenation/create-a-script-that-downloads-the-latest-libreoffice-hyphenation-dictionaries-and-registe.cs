using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for all output artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Create a deterministic local hyphenation dictionary.
        // The dictionary follows the OpenOffice format: language code,
        // number of patterns, then the pattern lines.
        // This minimal example is sufficient for registration and
        // demonstrates the API usage without external network calls.
        // -----------------------------------------------------------------
        string dictDir = Path.Combine(artifactsDir, "Dictionaries");
        Directory.CreateDirectory(dictDir);

        string germanDictPath = Path.Combine(dictDir, "de-CH.dic");
        // Very small dictionary – it contains a single trivial pattern.
        // In a real scenario you would use the full LibreOffice dictionary.
        string germanDictContent = @"de-CH
1
.1
";
        File.WriteAllText(germanDictPath, germanDictContent);

        // Register the German (Switzerland) hyphenation dictionary.
        Hyphenation.RegisterDictionary("de-CH", germanDictPath);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("de-CH"))
            throw new InvalidOperationException("Failed to register the de-CH hyphenation dictionary.");

        // -----------------------------------------------------------------
        // Build a sample document that will trigger hyphenation.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Set the font locale to German (Switzerland) so the registered dictionary is used.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Font.Size = 12;

        // Add a paragraph with long German text that can be hyphenated.
        builder.Writeln(
            "Die schnelle braune Füchsin springt über den faulen Hund. " +
            "Einige Wörter sind so lang, dass sie am Zeilenende getrennt werden müssen, " +
            "damit das Layout korrekt bleibt und keine übermäßigen Lücken entstehen.");

        // Save the document as PDF.
        string outputPath = Path.Combine(artifactsDir, "HyphenatedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF document was not created.", outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
