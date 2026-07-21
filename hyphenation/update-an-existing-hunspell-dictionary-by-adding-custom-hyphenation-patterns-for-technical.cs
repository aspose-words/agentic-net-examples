using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and output files.
        const string dictionaryPath = "hyph_en_US_custom.dic";
        const string outputPath = "HyphenatedCustom.pdf";

        // Create a minimal Hunspell hyphenation dictionary with custom patterns.
        // The first line must specify the encoding, e.g., UTF-8.
        // Subsequent lines contain words with hyphenation points marked by hyphens.
        string dictionaryContent =
            "UTF-8\n" +
            "microprocessor=micro-pro-cessor\n" +
            "hyperconvergence=hyper-con-ver-gence\n" +
            "quantumcomputing=quan-tum-com-put-ing\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the custom dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to increase the chance of line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set a very small hyphenation zone (must be > 0). Value is in 1/20 point.
        doc.HyphenationOptions.HyphenationZone = 1;

        // Write sample text containing the technical terms.
        builder.Font.Size = 12;
        builder.Writeln("The development of microprocessor technology has accelerated.");
        builder.Writeln("Recent advances in hyperconvergence are reshaping data centers.");
        builder.Writeln("Researchers explore quantumcomputing to solve complex problems.");

        // Save the document to PDF, which will apply hyphenation during layout.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Optional cleanup: unregister the dictionary if further processing is needed.
        // Hyphenation.UnregisterDictionary("en-US");
    }
}
