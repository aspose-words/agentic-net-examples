using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the custom hyphenation dictionary.
        const string dictionaryPath = "hyph_en_US_custom.dic";

        // Create a minimal Hunspell dictionary with custom patterns for technical terms.
        // The first line must specify the encoding.
        // Each subsequent line maps a word to its hyphenation pattern (hyphens separate syllables).
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "hyperparameter=hy-per-pa-ra-me-ter\n" +
            "multithreading=mul-ti-thread-ing\n" +
            "asynchronouscommunication=as-ync-ro-nous-com-mu-ni-ca-tion\n");

        // Register the dictionary for the English (US) locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the page layout to force line wrapping (narrow width).
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Insert a paragraph containing technical terminology that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("hyperparameter multithreading asynchronouscommunication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 360; // default zone

        // Save the document to PDF to visualize hyphenation.
        const string outputPath = "HyphenatedOutput.pdf";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
