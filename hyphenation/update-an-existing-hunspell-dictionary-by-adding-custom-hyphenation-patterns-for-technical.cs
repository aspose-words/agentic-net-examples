using System;
using System.IO;
using System.Linq;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(baseDir, "Data");
        Directory.CreateDirectory(dataDir);
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Hunspell hyphenation dictionary for English (US).
        // Format: first line = number of patterns, followed by pattern lines.
        string dictPath = Path.Combine(dataDir, "hyph_en_US.dic");
        var basePatterns = new[]
        {
            // Example generic pattern.
            "ab1c"
        };
        // Custom pattern for technical term "microprocessor".
        // The digit indicates a hyphenation point after the preceding characters.
        var customPatterns = new[]
        {
            "micro1processor"
        };
        var allPatterns = basePatterns.Concat(customPatterns).ToArray();
        var dictLines = new[] { allPatterns.Length.ToString() }.Concat(allPatterns);
        File.WriteAllLines(dictPath, dictLines);

        // Register the dictionary for the "en-US" locale.
        using (FileStream dictStream = new FileStream(dictPath, FileMode.Open, FileAccess.Read))
        {
            Hyphenation.RegisterDictionary("en-US", dictStream);
        }

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a document with narrow page width to force line breaks.
        Document doc = new Document();
        doc.FirstSection.PageSetup.PageWidth = 200; // points (~2.78 inches)
        doc.FirstSection.PageSetup.PageHeight = 842; // A4 height.

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // default

        // Add a paragraph containing a technical term that will be hyphenated using the custom pattern.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln(
            "The microprocessorarchitecture of modern CPUs is complex. " +
            "Hyphenation helps to keep the layout tidy when dealing with long technical terminology.");

        // Save the document to PDF to visualize hyphenation.
        string outputPath = Path.Combine(artifactsDir, "HyphenationCustomDictionary.pdf");
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF output was not created.", outputPath);
    }
}
