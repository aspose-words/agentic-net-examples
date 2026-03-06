using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;

namespace HyphenationExample
{
    class Program
    {
        static void Main()
        {
            // Folder paths – adjust as needed.
            string dataDir = @"C:\MyData\";
            string outputDir = @"C:\MyOutput\";

            // Register a warning collector to capture any issues during dictionary loading.
            WarningInfoCollection warnings = new WarningInfoCollection();
            Hyphenation.WarningCallback = warnings;

            // Register the English (US) hyphenation dictionary from a file.
            // This dictionary will be reused for any document that requires "en-US" hyphenation.
            Hyphenation.RegisterDictionary("en-US", Path.Combine(dataDir, "hyph_en_US.dic"));

            // Optional: verify that the dictionary is registered.
            if (!Hyphenation.IsDictionaryRegistered("en-US"))
                throw new InvalidOperationException("Failed to register the English hyphenation dictionary.");

            // Load the DOCM document that needs hyphenation.
            Document doc = new Document(Path.Combine(dataDir, "SampleDocument.docm"));

            // Enable automatic hyphenation for the document.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

            // Set a callback that will register additional dictionaries on demand.
            // For this example we only need the English dictionary, but the callback
            // demonstrates how to handle other languages (e.g., German).
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

            // Save the hyphenated document. The layout engine will invoke the callback
            // if it encounters a language without a registered dictionary.
            doc.Save(Path.Combine(outputDir, "SampleDocument_Hyphenated.pdf"));

            // Output any warnings that occurred during dictionary registration.
            Console.WriteLine($"Number of hyphenation warnings: {warnings.Count}");
        }
    }

    // Implements IHyphenationCallback to register dictionaries when requested.
    class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly string _basePath;
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string basePath)
        {
            _basePath = basePath;
            // Map language codes to dictionary file names.
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(_basePath, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(_basePath, "hyph_de_CH.dic") } // Example for German (Switzerland).
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

            // Register the dictionary if we have a known file for the language.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary available – the callback will not register anything.
            Console.WriteLine(", no dictionary file found.");
        }
    }
}
