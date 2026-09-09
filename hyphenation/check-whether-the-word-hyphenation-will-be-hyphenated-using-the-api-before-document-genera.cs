using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the local hyphenation dictionary.
        const string dictionaryPath = "hyph_en_US.dic";

        // Create a minimal dictionary that contains a hyphenation pattern for the word "hyphenation".
        // The first line must specify the encoding (e.g., UTF-8).
        // Subsequent lines define hyphenation patterns: word=pattern.
        File.WriteAllText(dictionaryPath, "UTF-8\nhyphenation=hy-phen-a-tion\n");

        // Check registration status before loading the dictionary.
        bool isRegisteredBefore = Hyphenation.IsDictionaryRegistered("en-US");
        Console.WriteLine($"Dictionary registered before loading: {isRegisteredBefore}");

        // Register the dictionary for the English (US) locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary is now registered.
        bool isRegisteredAfter = Hyphenation.IsDictionaryRegistered("en-US");
        Console.WriteLine($"Dictionary registered after loading: {isRegisteredAfter}");

        // If the dictionary is registered, the word "hyphenation" can be hyphenated.
        if (isRegisteredAfter)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure the page to be narrow so that hyphenation may occur.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Add a paragraph containing the target word.
            builder.Writeln("The process of hyphenation can affect the layout of a document. hyphenation");

            // Save the document to a PDF to force layout processing.
            const string outputPath = "HyphenationCheck.pdf";
            doc.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Document saved to {outputPath}");

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Expected PDF output was not created.");
        }
        else
        {
            Console.WriteLine("Hyphenation dictionary could not be registered; the word will not be hyphenated.");
        }
    }
}
