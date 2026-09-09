using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Directory to hold hyphenation dictionary files.
        string dictDir = "HyphenationDictionaries";
        Directory.CreateDirectory(dictDir);

        // Minimal valid dictionary contents for demonstration.
        var sampleDictionaries = new Dictionary<string, string>
        {
            { "en-US", "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" },
            { "de-CH", "UTF-8\ninternationalisierung=in-ter-na-tion-alisie-rung\n" }
        };

        // Create files and register each dictionary with Aspose.Words.
        foreach (var kvp in sampleDictionaries)
        {
            string filePath = Path.Combine(dictDir, $"hyph_{kvp.Key}.dic");
            File.WriteAllText(filePath, kvp.Value);
            Hyphenation.RegisterDictionary(kvp.Key, filePath);
        }

        // List all dictionary files found in the directory and display their language codes.
        Console.WriteLine("Available hyphenation dictionaries:");
        foreach (string filePath in Directory.GetFiles(dictDir, "*.dic"))
        {
            string fileName = Path.GetFileName(filePath);
            // Expected naming pattern: hyph_{languageCode}.dic
            string languageCode = fileName.StartsWith("hyph_") && fileName.EndsWith(".dic")
                ? fileName.Substring(5, fileName.Length - 5 - 4)
                : "Unknown";

            Console.WriteLine($"- {languageCode}");
        }
    }
}
