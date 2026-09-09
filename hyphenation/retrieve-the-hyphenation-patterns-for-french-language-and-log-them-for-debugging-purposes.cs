using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the temporary French hyphenation dictionary.
        const string dictionaryPath = "hyph_fr_FR.dic";

        // Create a minimal French hyphenation dictionary.
        // The first line must specify the encoding (UTF-8).
        // Subsequent lines contain word=hyphenation-pattern pairs.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinaire=ex-tra-or-di-nai-re\n" +
            "internationalisation=in-ter-na-tio-na-li-sa-tion\n" +
            "communication=co-mmu-ni-ca-tion\n";

        // Write the dictionary file to the local file system.
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary for the French locale (fr-FR).
        // The Hyphenation class resides directly in the Aspose.Words namespace,
        // so we can reference it without an additional using directive.
        Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Verify that the dictionary was successfully registered.
        if (!Hyphenation.IsDictionaryRegistered("fr-FR"))
            throw new InvalidOperationException("Failed to register the French hyphenation dictionary.");

        // Retrieve and log the hyphenation patterns for debugging.
        Console.WriteLine("Hyphenation patterns for French (fr-FR):");
        foreach (string line in File.ReadAllLines(dictionaryPath))
        {
            // Skip the encoding header line.
            if (line.StartsWith("UTF-8", StringComparison.OrdinalIgnoreCase))
                continue;

            Console.WriteLine(line);
        }

        // Clean up the temporary dictionary file.
        File.Delete(dictionaryPath);
    }
}
