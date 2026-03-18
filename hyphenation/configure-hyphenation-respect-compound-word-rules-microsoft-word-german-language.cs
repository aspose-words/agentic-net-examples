using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Paths to input and output folders (adjust as needed).
        string dataDir = @"C:\Data\";
        string artifactsDir = @"C:\Artifacts\";

        // Register the German hyphenation dictionary (de-CH) from a file.
        Hyphenation.RegisterDictionary("de-CH", Path.Combine(dataDir, "hyph_de_CH.dic"));

        // Optional: set a callback to load dictionaries on demand.
        Hyphenation.Callback = new HyphenationCallback(dataDir);

        // Load a German language document.
        Document doc = new Document(Path.Combine(dataDir, "GermanText.docx"));

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Keep default behavior for capitalized words.
        doc.HyphenationOptions.HyphenateCaps = true;
        // Limit consecutive hyphenated lines (optional).
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // Set hyphenation zone (distance from right margin where hyphenation is suppressed).
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 1440 points)

        // Save the hyphenated document.
        doc.Save(Path.Combine(artifactsDir, "GermanHyphenated.docx"));
    }

    // Callback that registers hyphenation dictionaries when requested by the layout engine.
    private class HyphenationCallback : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public HyphenationCallback(string basePath)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "de-CH", Path.Combine(basePath, "hyph_de_CH.dic") },
                { "en-US", Path.Combine(basePath, "hyph_en_US.dic") }
            };
        }

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a matching file.
            if (_dictionaryFiles.TryGetValue(language, out string filePath))
                Hyphenation.RegisterDictionary(language, filePath);
        }
    }
}
