using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

namespace HyphenationExample
{
    // Custom callback that registers hyphenation dictionaries on demand.
    class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string dictionariesFolder)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(dictionariesFolder, "hyph_de_CH.dic") }
                // Add more language‑file mappings as needed.
            };
        }

        public void RequestDictionary(string language)
        {
            Console.Write($"Hyphenation dictionary requested: {language}");

            // If already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine(", already registered.");
                return;
            }

            // Register the dictionary if we know its location.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary available – register a null dictionary to suppress further callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
            Console.WriteLine(", no dictionary found – null dictionary registered.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder that contains the hyphenation dictionary files.
            string myDir = @"C:\HyphenationDictionaries";

            // Create a new blank document.
            Document doc = new Document();

            // Build content with a locale that requires hyphenation (German - Switzerland).
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Size = 24;
            builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
            builder.Writeln("Dies ist ein Beispieltext, der die Silbentrennung demonstriert, wenn das Wort über mehrere Zeilen hinweg fortgesetzt wird.");

            // Register the German (Switzerland) hyphenation dictionary from file.
            string germanDictionaryPath = Path.Combine(myDir, "hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", germanDictionaryPath);

            // Enable automatic hyphenation for the whole document.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from right margin.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenateCaps = true;

            // Set a callback to handle any other languages that might be encountered.
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister(myDir);

            // Save the document.
            string outputPath = @"C:\Output\HyphenatedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
