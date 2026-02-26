using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Directory that contains the DOCM file and the OpenOffice hyphenation dictionaries.
        string dataDir = @"C:\Docs\";

        // Input DOCM document.
        string inputPath = Path.Combine(dataDir, "Sample.docm");

        // Output PDF where hyphenation will be applied.
        string outputPath = Path.Combine(dataDir, "Sample_Hyphenated.pdf");

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Turn on automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Register a callback that will supply OpenOffice hyphenation dictionaries when needed.
        Hyphenation.Callback = new OpenOfficeDictionaryCallback(dataDir);

        // Save the document to PDF. During layout the callback will be invoked,
        // the required dictionaries will be registered, and hyphenation will be applied.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}

// Callback implementation that registers hyphenation dictionaries from OpenOffice .dic files.
class OpenOfficeDictionaryCallback : IHyphenationCallback
{
    // Maps language codes (e.g., "en-US") to the corresponding dictionary file path.
    private readonly Dictionary<string, string> _dictionaryFiles;

    public OpenOfficeDictionaryCallback(string baseDirectory)
    {
        // Populate the map with the languages you expect to encounter.
        // Add or remove entries as needed.
        _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en-US", Path.Combine(baseDirectory, "hyph_en_US.dic") },
            { "de-CH", Path.Combine(baseDirectory, "hyph_de_CH.dic") },
            // Example for additional language:
            // { "fr-FR", Path.Combine(baseDirectory, "hyph_fr_FR.dic") }
        };
    }

    // This method is called by Aspose.Words when a dictionary for a language is not yet registered.
    public void RequestDictionary(string language)
    {
        // If the dictionary is already registered, do nothing.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Try to find a dictionary file for the requested language.
        if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
        {
            // Register the dictionary from the file.
            Hyphenation.RegisterDictionary(language, filePath);
        }
        else
        {
            // No dictionary available – register a null dictionary to prevent repeated callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
