using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the French hyphenation dictionary.
        const string dictionaryPath = "hyph_fr_FR.dic";

        // Create a minimal French hyphenation dictionary.
        string[] dictionaryContent =
        {
            "UTF-8",
            "bonjour=bon-jour",
            "extraordinaire=ex-tra-or-di-nai-re"
        };
        File.WriteAllLines(dictionaryPath, dictionaryContent);

        // Verify the dictionary file was created.
        if (!File.Exists(dictionaryPath))
            throw new InvalidOperationException("Failed to create the hyphenation dictionary file.");

        // Register the dictionary with Aspose.Words.
        Aspose.Words.Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Retrieve and log the hyphenation patterns for French.
        string[] allLines = File.ReadAllLines(dictionaryPath);
        var patterns = allLines.Skip(1); // Skip the encoding line.

        Console.WriteLine("Hyphenation patterns for French (fr-FR):");
        foreach (string pattern in patterns)
        {
            Console.WriteLine(pattern);
        }
    }
}
