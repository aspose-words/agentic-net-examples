using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create sample hyphenation dictionary files in the current directory.
        const string enDictionaryPath = "hyph_en_US.dic";
        const string frDictionaryPath = "hyph_fr_FR.dic";

        File.WriteAllText(enDictionaryPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        File.WriteAllText(frDictionaryPath,
            "UTF-8\nbonjour=bon-jour\n");

        // Register the dictionaries with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", enDictionaryPath);
        Hyphenation.RegisterDictionary("fr-FR", frDictionaryPath);

        // Retrieve and display all hyphenation dictionaries that exist in the current folder.
        string[] dictionaryFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "hyph_*.dic");
        var languageCodes = new List<string>();

        foreach (string filePath in dictionaryFiles)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath); // e.g., hyph_en_US
            string[] parts = fileName.Split('_', 2);
            if (parts.Length == 2)
            {
                // Convert underscore separator to hyphen as used by Aspose.Words.
                string code = parts[1].Replace('_', '-');
                languageCodes.Add(code);
            }
        }

        Console.WriteLine("Available hyphenation dictionaries:");
        foreach (string code in languageCodes)
        {
            Console.WriteLine(code);
        }

        // Clean up the temporary dictionary files (optional).
        try
        {
            File.Delete(enDictionaryPath);
            File.Delete(frDictionaryPath);
        }
        catch
        {
            // Ignored – cleanup failure should not affect the example outcome.
        }
    }
}
