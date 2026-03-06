using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace HyphenationDemo
{
    // Implements the callback that registers hyphenation dictionaries on demand.
    public class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        // Map language codes to dictionary file paths.
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string dictionariesFolder)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(dictionariesFolder, "hyph_de_CH.dic") }
                // Add more language‑dictionary pairs as needed.
            };
        }

        // This method is called by Aspose.Words when a dictionary for a language is missing.
        public void RequestDictionary(string language)
        {
            Console.Write($"Hyphenation dictionary requested: {language}");

            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine(", already registered.");
                return;
            }

            // If we know the file for the requested language, register it.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary available – optionally register a null dictionary to suppress further callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
            Console.WriteLine(", not found. Registered null dictionary to avoid repeated requests.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder where input documents, dictionaries and output files reside.
            string dataDir = @"C:\Data";
            string dictionariesDir = Path.Combine(dataDir, "Dictionaries");
            string inputDocPath = Path.Combine(dataDir, "German text.docx");
            string outputPdfPath = Path.Combine(dataDir, "HyphenatedOutput.pdf");

            // Optional: capture hyphenation warnings.
            WarningInfoCollection warnings = new WarningInfoCollection();
            Hyphenation.WarningCallback = warnings;

            // Pre‑register an English dictionary (optional, demonstrates direct registration).
            using (Stream enStream = new FileStream(Path.Combine(dictionariesDir, "hyph_en_US.dic"), FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", enStream);
            }

            // Load the document (uses the provided load rule).
            Document doc = new Document(inputDocPath);

            // Assign the callback that will load dictionaries on demand.
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dictionariesDir);

            // Enable automatic hyphenation for the document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Save the document (uses the provided save rule).
            doc.Save(outputPdfPath, SaveFormat.Pdf);

            // Output any hyphenation warnings that occurred.
            if (warnings.Count > 0)
            {
                Console.WriteLine("Hyphenation warnings:");
                foreach (var warning in warnings)
                {
                    Console.WriteLine($"- {warning.Description}");
                }
            }
        }
    }
}
