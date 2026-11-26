using System;
using System.IO;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Create sample hyphenation dictionaries locally.
        var dictionaries = new[]
        {
            new
            {
                LanguageCode = "en-US",
                FileName = "hyph_en_US.dic",
                Content = "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n"
            },
            new
            {
                LanguageCode = "de-DE",
                FileName = "hyph_de_DE.dic",
                Content = "UTF-8\nBeispiel=Bei-spiel\n"
            }
        };

        foreach (var dict in dictionaries)
        {
            // Ensure the dictionary file exists before registration.
            File.WriteAllText(dict.FileName, dict.Content);
            Aspose.Words.Hyphenation.RegisterDictionary(dict.LanguageCode, dict.FileName);
        }

        // List all hyphenation dictionaries present in the current directory.
        var dicFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "hyph_*.dic");

        Console.WriteLine("Available hyphenation dictionaries (language codes):");
        foreach (var filePath in dicFiles)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath); // e.g., hyph_en_US
            var parts = fileName.Split('_');
            if (parts.Length >= 2)
            {
                // Combine remaining parts with hyphen to form language code (en-US, de-DE, etc.).
                var languageCode = string.Join("-", parts.Skip(1));
                Console.WriteLine(languageCode);
            }
        }
    }
}
