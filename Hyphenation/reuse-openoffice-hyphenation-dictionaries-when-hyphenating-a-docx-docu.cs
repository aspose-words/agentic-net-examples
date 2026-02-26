using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words; // Hyphenation classes are in the Aspose.Words namespace for recent versions.

class Program
{
    static void Main()
    {
        // Paths to input and output folders – adjust to your environment.
        string MyDir = @"C:\Data\";
        string ArtifactsDir = @"C:\Output\";

        // Optionally pre‑register a dictionary that is always needed.
        Hyphenation.RegisterDictionary("en-US", Path.Combine(MyDir, "hyph_en_US.dic"));

        // Set a callback that will load dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(MyDir);

        // Load the DOCX document that requires hyphenation.
        Document doc = new Document(Path.Combine(MyDir, "German text.docx"));

        // Turn on automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the hyphenated result (PDF format demonstrates hyphenation).
        doc.Save(Path.Combine(ArtifactsDir, "GermanHyphenated.pdf"));
    }
}

// Implements IHyphenationCallback to supply OpenOffice .dic files when requested.
class CustomHyphenationDictionaryRegister : IHyphenationCallback
{
    private readonly Dictionary<string, string> _dictionaryFiles;

    public CustomHyphenationDictionaryRegister(string baseDir)
    {
        // Map language codes to the corresponding OpenOffice dictionary files.
        _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en-US", Path.Combine(baseDir, "hyph_en_US.dic") },
            { "de-CH", Path.Combine(baseDir, "hyph_de_CH.dic") }
            // Add additional mappings as required.
        };
    }

    public void RequestDictionary(string language)
    {
        // If the dictionary is already registered, no action is needed.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Register the dictionary if we have a file for the requested language.
        if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
        {
            Hyphenation.RegisterDictionary(language, filePath);
        }
        else
        {
            // Register a null dictionary to prevent repeated callbacks for unknown languages.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
