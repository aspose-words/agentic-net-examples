using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the temporary French hyphenation dictionary.
        const string dictionaryPath = "hyph_fr_FR.dic";

        // Minimal dictionary content: encoding header + a few word patterns.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinaire=ex-tra-or-di-nai-re\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary file to the local file system.
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the French hyphenation dictionary.
        Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("fr-FR"))
            throw new InvalidOperationException("Failed to register the French hyphenation dictionary.");

        // Log the dictionary entries (skip the encoding header).
        Console.WriteLine("French hyphenation patterns loaded from dictionary:");
        foreach (string line in File.ReadLines(dictionaryPath))
        {
            if (line.StartsWith("UTF-", StringComparison.OrdinalIgnoreCase))
                continue;

            Console.WriteLine(line);
        }

        // Create a simple document containing French words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("fr-FR").LCID;
        builder.Writeln("extraordinaire communication");

        // Enable automatic hyphenation so the dictionary is used.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document.
        const string outputPath = "FrenchHyphenated.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
