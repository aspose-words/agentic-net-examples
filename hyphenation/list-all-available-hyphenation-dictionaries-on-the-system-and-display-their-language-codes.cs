using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Directory to store sample hyphenation dictionaries.
        const string dictDirectory = "HyphenationDictionaries";
        Directory.CreateDirectory(dictDirectory);

        // Create a minimal English (US) hyphenation dictionary.
        string enDictPath = Path.Combine(dictDirectory, "hyph_en_US.dic");
        File.WriteAllText(enDictPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the English dictionary.
        Hyphenation.RegisterDictionary("en-US", enDictPath);

        // Create a minimal German (Switzerland) hyphenation dictionary.
        string deDictPath = Path.Combine(dictDirectory, "hyph_de_CH.dic");
        File.WriteAllText(deDictPath,
            "UTF-8\nkommunikation=kom-mu-ni-ka-tion\ninternationalisierung=inter-na-tion-ali-sier-ung\n");

        // Register the German dictionary.
        Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // List all dictionary files in the directory and display their language codes.
        Console.WriteLine("Available hyphenation dictionaries:");
        foreach (string filePath in Directory.GetFiles(dictDirectory, "hyph_*.dic"))
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath); // e.g., hyph_en_US
            if (fileName.Length <= "hyph_".Length)
                continue;

            // Extract the language part after the "hyph_" prefix.
            string languagePart = fileName.Substring("hyph_".Length); // e.g., en_US
            // Convert underscore to hyphen to match the culture name format.
            string languageCode = languagePart.Replace('_', '-');

            // Verify that the dictionary is registered.
            bool isRegistered = Hyphenation.IsDictionaryRegistered(languageCode);
            Console.WriteLine($"- File: {Path.GetFileName(filePath)} | Language code: {languageCode} | Registered: {isRegistered}");
        }
    }
}
