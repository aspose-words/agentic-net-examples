using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Folder that contains the source document and the OpenOffice hyphenation dictionaries.
        string dataDir = @"C:\Data\";

        // Load the DOC/DOCX document that needs hyphenation.
        Document doc = new Document(Path.Combine(dataDir, "German text.docx"));

        // Enable automatic hyphenation and configure optional parameters.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 20 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Register a callback that will supply dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

        // Save the document; during layout the callback will load the required dictionaries.
        doc.Save(Path.Combine(dataDir, "Hyphenated.pdf"));
    }
}

// Callback implementation that maps language codes to OpenOffice dictionary files.
class CustomHyphenationDictionaryRegister : IHyphenationCallback
{
    private readonly Dictionary<string, string> _dictionaryFiles;

    public CustomHyphenationDictionaryRegister(string baseDirectory)
    {
        // Populate the map with the languages you intend to support.
        _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en-US", Path.Combine(baseDirectory, "hyph_en_US.dic") },
            { "de-CH", Path.Combine(baseDirectory, "hyph_de_CH.dic") }
            // Add more entries as needed, e.g. { "fr-FR", Path.Combine(baseDirectory, "hyph_fr_FR.dic") }
        };
    }

    // This method is called by Aspose.Words when a dictionary for a language is missing.
    public void RequestDictionary(string language)
    {
        // If the dictionary is already registered, nothing to do.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Try to locate a dictionary file for the requested language.
        if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
        {
            // Register the dictionary from the file.
            Hyphenation.RegisterDictionary(language, filePath);
        }
        else
        {
            // No dictionary available – optionally register a null dictionary to suppress further callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
