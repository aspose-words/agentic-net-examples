using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Paths to the input files and output folder.
        string myDir = @"C:\MyDir\";
        string artifactsDir = @"C:\Artifacts\";

        // Register an English dictionary upfront (optional, demonstrates RegisterDictionary overload with Stream).
        using (Stream enStream = new FileStream(Path.Combine(myDir, "hyph_en_US.dic"), FileMode.Open))
        {
            Hyphenation.RegisterDictionary("en-US", enStream);
        }

        // Set a callback that will load dictionaries on demand when the layout engine needs them.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(myDir);

        // Load a DOC document that contains text in a language for which we have a dictionary (e.g., German - de-CH).
        Document doc = new Document(Path.Combine(myDir, "German text.doc"));

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document (PDF format forces layout processing and thus hyphenation).
        doc.Save(Path.Combine(artifactsDir, "Hyphenated.pdf"));
    }

    // Implementation of IHyphenationCallback that registers dictionaries from local files.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string baseDir)
        {
            // Map language codes to dictionary file paths.
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(baseDir, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(baseDir, "hyph_de_CH.dic") }
            };
        }

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, nothing to do.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a matching file.
            if (_dictionaryFiles.TryGetValue(language, out string filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                return;
            }

            // No dictionary available – register a null dictionary to suppress further callbacks for this language.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
