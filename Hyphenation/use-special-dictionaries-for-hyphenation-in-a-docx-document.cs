using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class HyphenationExample
{
    // Paths to the input document, hyphenation dictionaries and output folder.
    private const string MyDir = @"C:\MyDir\";
    private const string ArtifactsDir = @"C:\Artifacts\";

    static void Main()
    {
        // Register a dictionary directly (optional, can also rely on the callback).
        Hyphenation.RegisterDictionary("en-US", Path.Combine(MyDir, "hyph_en_US.dic"));

        // Set up a callback that will register dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

        // Load a document that contains text in a language for which we have a dictionary.
        Document doc = new Document(Path.Combine(MyDir, "German text.docx"));

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Save the hyphenated document.
        doc.Save(Path.Combine(ArtifactsDir, "HyphenatedDocument.docx"));
    }

    // Callback that registers hyphenation dictionaries when the layout engine requests them.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister()
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(MyDir, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(MyDir, "hyph_de_CH.dic") },
                // Add more language‑code / file‑path pairs as needed.
            };
        }

        public void RequestDictionary(string language)
        {
            Console.Write($"Hyphenation dictionary requested: {language}");

            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine(", already registered.");
                return;
            }

            // Register the dictionary if we have a matching file.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary known for this language.
            Console.WriteLine(", no dictionary file known.");
        }
    }
}
