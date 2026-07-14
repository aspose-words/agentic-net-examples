using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder that will hold the hyphenation dictionary files.
        const string dictFolder = "HyphenationDictionaries";
        Directory.CreateDirectory(dictFolder);

        // Define a few sample language codes and create minimal dictionary files for them.
        string[] languageCodes = { "en-US", "de-CH", "fr-FR" };
        foreach (string lang in languageCodes)
        {
            // Convert the language code to a file name, e.g. en-US -> hyph_en_US.dic
            string fileName = Path.Combine(dictFolder, $"hyph_{lang.Replace("-", "_")}.dic");

            // Minimal valid dictionary content (UTF‑8 header plus a single pattern).
            // The actual patterns are not important for this example.
            string content = "UTF-8\nsample=sa-mple\n";
            File.WriteAllText(fileName, content);

            // Register the dictionary with Aspose.Words.
            Hyphenation.RegisterDictionary(lang, fileName);
        }

        // List all dictionary files that exist in the folder and display their language codes.
        Console.WriteLine("Available hyphenation dictionaries:");
        foreach (string filePath in Directory.GetFiles(dictFolder, "*.dic"))
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath); // e.g. hyph_en_US
            if (fileName.StartsWith("hyph_", StringComparison.OrdinalIgnoreCase))
            {
                // Extract the language part and convert back to the standard format (en_US -> en-US).
                string langPart = fileName.Substring("hyph_".Length).Replace('_', '-');
                Console.WriteLine($"- {langPart}");
            }
        }

        // Verify that each dictionary is indeed registered.
        foreach (string lang in languageCodes)
        {
            if (!Hyphenation.IsDictionaryRegistered(lang))
                throw new InvalidOperationException($"Hyphenation dictionary for '{lang}' was not registered.");
        }
    }
}
