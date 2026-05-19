using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the path for the French hyphenation dictionary.
        string dictionaryPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_fr_FR.dic");

        // Create a minimal valid French hyphenation dictionary.
        // The first line must specify the encoding (UTF-8), followed by word=pattern lines.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinaire=ex-tra-or-di-nai-re\n" +
            "internationalisation=in-ter-na-tio-na-li-sa-tion\n" +
            "communication=co-mmu-ni-ca-tion\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the French dictionary with the language code "fr-FR".
        Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Verify that the dictionary was successfully registered.
        if (!Hyphenation.IsDictionaryRegistered("fr-FR"))
            throw new InvalidOperationException("Failed to register the French hyphenation dictionary.");

        // Retrieve the patterns by reading the dictionary file (for debugging purposes).
        string loadedPatterns = File.ReadAllText(dictionaryPath);
        Console.WriteLine("French hyphenation patterns:");
        Console.WriteLine(loadedPatterns);

        // Optional: demonstrate that the dictionary works by creating a document with French text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinaire internationalisation communication");
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // Increase zone to encourage hyphenation.
        doc.Save("FrenchHyphenated.pdf");

        // Validate that the output file was created.
        if (!File.Exists("FrenchHyphenated.pdf"))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
