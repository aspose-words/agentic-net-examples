using System;
using System.IO;
using Aspose.Words;

public class ListHyphenationDictionaries
{
    public static void Main()
    {
        // Create a folder to hold sample hyphenation dictionaries.
        string dictFolder = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDictionaries");
        Directory.CreateDirectory(dictFolder);

        // Sample dictionary for English (US).
        string enDictPath = Path.Combine(dictFolder, "hyph_en_US.dic");
        File.WriteAllText(enDictPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Sample dictionary for German (Switzerland).
        string deDictPath = Path.Combine(dictFolder, "hyph_de_CH.dic");
        File.WriteAllText(deDictPath,
            "UTF-8\nkommunikation=kom-mu-ni-ka-tion\n");

        // Find all dictionary files in the folder.
        string[] dictFiles = Directory.GetFiles(dictFolder, "*.dic");
        if (dictFiles.Length == 0)
            throw new InvalidOperationException("No hyphenation dictionary files were found.");

        Console.WriteLine("Available hyphenation dictionaries:");
        foreach (string filePath in dictFiles)
        {
            // Extract language code from file name: hyph_{code}.dic -> {code}
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            if (!fileName.StartsWith("hyph_"))
                continue; // Skip files that do not follow the naming convention.

            string languageCode = fileName.Substring(5).Replace('_', '-'); // e.g., en_US -> en-US

            // Register the dictionary for the extracted language.
            Hyphenation.RegisterDictionary(languageCode, filePath);

            // Verify registration.
            if (!Hyphenation.IsDictionaryRegistered(languageCode))
                throw new InvalidOperationException($"Failed to register dictionary for language '{languageCode}'.");

            // Output the language code.
            Console.WriteLine($"- {languageCode}");
        }
    }
}
