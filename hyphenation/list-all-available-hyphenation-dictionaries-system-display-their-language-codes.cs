using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;

// Define the folder that contains hyphenation dictionary files.
// Adjust this path to point to the actual location of the dictionaries on your system.
string hyphenationFolder = @"C:\HyphenationDictionaries";

// Verify that the folder exists.
if (!Directory.Exists(hyphenationFolder))
{
    Console.WriteLine($"Folder not found: {hyphenationFolder}");
    return;
}

// Pattern to extract the language code from a typical dictionary filename, e.g. "hyph_en_US.dic".
Regex langRegex = new Regex(@"hyph_(?<lang>[\w-]+)\.dic$", RegexOptions.IgnoreCase);

foreach (string filePath in Directory.GetFiles(hyphenationFolder, "*.dic"))
{
    string fileName = Path.GetFileName(filePath);
    Match match = langRegex.Match(fileName);
    if (!match.Success)
    {
        // Skip files that do not follow the expected naming convention.
        continue;
    }

    string languageCode = match.Groups["lang"].Value.Replace('_', '-'); // Convert "en_US" to "en-US".

    // Register the dictionary if it hasn't been registered yet.
    if (!Hyphenation.IsDictionaryRegistered(languageCode))
    {
        try
        {
            Hyphenation.RegisterDictionary(languageCode, filePath);
            Console.WriteLine($"Registered hyphenation dictionary: {languageCode}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to register '{fileName}': {ex.Message}");
        }
    }
    else
    {
        Console.WriteLine($"Dictionary already registered: {languageCode}");
    }
}

// After registration, list all language codes that are currently registered.
Console.WriteLine("\nAvailable hyphenation dictionaries (language codes):");
foreach (string filePath in Directory.GetFiles(hyphenationFolder, "*.dic"))
{
    string fileName = Path.GetFileName(filePath);
    Match match = langRegex.Match(fileName);
    if (!match.Success) continue;

    string languageCode = match.Groups["lang"].Value.Replace('_', '-');
    if (Hyphenation.IsDictionaryRegistered(languageCode))
    {
        Console.WriteLine($"- {languageCode}");
    }
}
