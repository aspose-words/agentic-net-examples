using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // List of language codes to query.
        List<string> languages = new List<string>
        {
            "en-US", // English (United States)
            "de-CH", // German (Switzerland)
            "fr-FR", // French (France)
            "es-ES", // Spanish (Spain)
            "ja-JP"  // Japanese
        };

        Console.WriteLine("Registered hyphenation dictionaries on this system:");
        foreach (string lang in languages)
        {
            bool isRegistered = Hyphenation.IsDictionaryRegistered(lang);
            Console.WriteLine($"{lang}: {(isRegistered ? "Registered" : "Not registered")}");
        }

        // Example of registering a minimal dictionary for demonstration.
        const string sampleLang = "en-US";
        const string sampleDictFile = "sample_hyph_en_US.dic";

        // Create a very small, valid dictionary file if it does not already exist.
        if (!File.Exists(sampleDictFile))
        {
            string[] lines =
            {
                "SET UTF-8",
                "% Minimal hyphenation dictionary for demonstration",
                "1",
                "a1",
                "b1"
            };
            File.WriteAllLines(sampleDictFile, lines);
        }

        // Register the sample dictionary.
        Hyphenation.RegisterDictionary(sampleLang, sampleDictFile);
        Console.WriteLine($"\nAfter registering '{sampleDictFile}':");
        Console.WriteLine($"{sampleLang}: {(Hyphenation.IsDictionaryRegistered(sampleLang) ? "Registered" : "Not registered")}");

        // Clean up: unregister the sample dictionary.
        Hyphenation.UnregisterDictionary(sampleLang);
    }
}
