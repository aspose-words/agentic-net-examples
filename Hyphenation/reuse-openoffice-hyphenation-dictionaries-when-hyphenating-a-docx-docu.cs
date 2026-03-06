using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Paths to the input DOCX, output PDF and the folder that contains OpenOffice hyphenation dictionaries.
        string inputDocx = @"C:\Docs\German text.docx";
        string outputPdf = @"C:\Docs\Hyphenated output.pdf";
        string dictionariesFolder = @"C:\HyphenationDictionaries\";

        // Register any dictionaries that are known in advance (optional).
        // Example: register English dictionary immediately.
        Hyphenation.RegisterDictionary("en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic"));

        // Set up a callback that will be invoked when the layout engine needs a dictionary for a language
        // that has not yet been registered. The callback will look for a matching .dic file in the folder.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dictionariesFolder);

        // Load the DOCX document.
        Document doc = new Document(inputDocx);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to a fixed‑page format (PDF) so that hyphenation is applied.
        doc.Save(outputPdf);
    }
}

// Callback implementation that registers dictionaries from a specified folder on demand.
class CustomHyphenationDictionaryRegister : IHyphenationCallback
{
    private readonly string _folderPath;
    private readonly Dictionary<string, string> _cachedPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public CustomHyphenationDictionaryRegister(string folderPath)
    {
        _folderPath = folderPath;
    }

    public void RequestDictionary(string language)
    {
        // If the dictionary is already registered, nothing to do.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Try to find a dictionary file for the requested language.
        string dicFile = GetDictionaryPath(language);
        if (dicFile != null && File.Exists(dicFile))
        {
            // Register the dictionary from the file.
            Hyphenation.RegisterDictionary(language, dicFile);
        }
        else
        {
            // No dictionary found – register a null dictionary to suppress further callbacks for this language.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }

    // Returns the full path to a .dic file that matches the language code, e.g. "de-CH" -> "hyph_de_CH.dic".
    private string GetDictionaryPath(string language)
    {
        if (_cachedPaths.TryGetValue(language, out string cached))
            return cached;

        // Build expected file name pattern.
        string fileName = $"hyph_{language.Replace('-', '_')}.dic";
        string fullPath = Path.Combine(_folderPath, fileName);

        // Cache the result (null if not found) for future calls.
        _cachedPaths[language] = File.Exists(fullPath) ? fullPath : null;
        return _cachedPaths[language];
    }
}
