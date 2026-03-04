using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Directories that contain the source DOCM and the OpenOffice hyphenation dictionaries.
        string dataDir   = @"C:\Data\";
        string outputDir = @"C:\Output\";

        // Load the DOCM document.
        Document doc = new Document(Path.Combine(dataDir, "Sample.docm"));

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Register a callback that will load the required OpenOffice hyphenation dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

        // Save the hyphenated document (PDF format is used here, but any format can be chosen).
        doc.Save(Path.Combine(outputDir, "Sample_Hyphenated.pdf"));
    }
}

// Callback that maps language codes to dictionary files and registers them when requested.
class CustomHyphenationDictionaryRegister : IHyphenationCallback
{
    private readonly Dictionary<string, string> _dictionaryFiles;

    public CustomHyphenationDictionaryRegister(string dictionariesFolder)
    {
        // Populate the map with the languages you need. Add more entries as required.
        _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic") },
            { "de-CH", Path.Combine(dictionariesFolder, "hyph_de_CH.dic") },
            // Example for another language:
            // { "fr-FR", Path.Combine(dictionariesFolder, "hyph_fr_FR.dic") }
        };
    }

    public void RequestDictionary(string language)
    {
        // If the dictionary is already registered, nothing more to do.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Register the dictionary if we have a matching file.
        if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
        {
            Hyphenation.RegisterDictionary(language, filePath);
        }
        // If no dictionary is known for the requested language, the callback simply does nothing.
        // The layout engine will then proceed without hyphenation for that language.
    }
}
