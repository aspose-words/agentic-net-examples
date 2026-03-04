using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Directories – adjust to your environment
        string MyDir = @"C:\Docs\";
        string ArtifactsDir = @"C:\Output\";

        // Register hyphenation dictionaries that will be used directly
        Hyphenation.RegisterDictionary("en-US", Path.Combine(MyDir, "hyph_en_US.dic"));
        Hyphenation.RegisterDictionary("de-CH", Path.Combine(MyDir, "hyph_de_CH.dic"));

        // Set a callback to register dictionaries on demand for other languages
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(MyDir);

        // Load a document that contains German (de-CH) text
        Document doc = new Document(Path.Combine(MyDir, "German text.docx"));

        // Enable automatic hyphenation for the whole document
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 20 = 36 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Suppress hyphenation for the first paragraph as an example
        doc.FirstSection.Body.FirstParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the hyphenated document
        doc.Save(Path.Combine(ArtifactsDir, "HyphenatedDocument.docx"));
    }

    // Callback that registers a hyphenation dictionary when the layout engine requests it
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string basePath)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(basePath, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(basePath, "hyph_de_CH.dic") }
                // Add additional language‑to‑file mappings here if needed
            };
        }

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, do nothing
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a file for the requested language
            if (_dictionaryFiles.TryGetValue(language, out string fileName))
                Hyphenation.RegisterDictionary(language, fileName);
        }
    }
}
