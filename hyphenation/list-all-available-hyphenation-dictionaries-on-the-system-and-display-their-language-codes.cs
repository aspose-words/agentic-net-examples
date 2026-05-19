using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define language codes for which we will create dummy hyphenation dictionaries.
        string[] languageCodes = { "en-US", "de-CH", "fr-FR", "es-ES" };

        // Create and register a minimal dictionary file for each language.
        foreach (string lang in languageCodes)
        {
            // Build a deterministic file name based on the language code.
            string fileName = $"hyph_{lang.Replace("-", "_")}.dic";

            // Write a minimal valid dictionary content. The first line must be the encoding.
            // No actual hyphenation patterns are required for this demonstration.
            File.WriteAllText(fileName, "UTF-8\n");

            // Register the dictionary with Aspose.Words.
            Hyphenation.RegisterDictionary(lang, fileName);
        }

        // List all language codes that have a registered hyphenation dictionary.
        Console.WriteLine("Registered hyphenation dictionaries:");
        foreach (string lang in languageCodes)
        {
            if (Hyphenation.IsDictionaryRegistered(lang))
            {
                Console.WriteLine($"- {lang}");
            }
        }

        // Optional cleanup: delete the temporary dictionary files.
        foreach (string lang in languageCodes)
        {
            string fileName = $"hyph_{lang.Replace("-", "_")}.dic";
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }
}
