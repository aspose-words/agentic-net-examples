using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace HyphenationExample
{
    // Implements the callback that Aspose.Words invokes when it needs a hyphenation dictionary.
    class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        // Maps language codes to the corresponding OpenOffice hyphenation dictionary files.
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string dictionariesFolder)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(dictionariesFolder, "hyph_de_CH.dic") },
                // Add more language mappings here as needed.
            };
        }

        // Called by Aspose.Words when a dictionary for the specified language is not yet registered.
        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, nothing to do.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Try to find a matching file and register it.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                return;
            }

            // No dictionary available – register a null dictionary to suppress further callbacks for this language.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder that contains the input document and the OpenOffice hyphenation dictionaries.
            string dataDir = @"C:\Data";

            // Pre‑register a dictionary that we know will be needed immediately.
            Hyphenation.RegisterDictionary("en-US", Path.Combine(dataDir, "hyph_en_US.dic"));

            // Set the callback so that additional dictionaries are loaded on demand.
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

            // Load the DOCX document that requires hyphenation.
            Document doc = new Document(Path.Combine(dataDir, "Input.docx"));

            // Enable automatic hyphenation for the whole document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: tweak hyphenation settings.
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenateCaps = true;

            // Save the hyphenated document. The format is inferred from the file extension.
            doc.Save(Path.Combine(dataDir, "Output.pdf"));
        }
    }
}
