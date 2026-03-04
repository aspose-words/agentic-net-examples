using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace HyphenationDemo
{
    // Implements the callback that registers hyphenation dictionaries on demand.
    public class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        // Mapping of language codes to dictionary file paths.
        private readonly Dictionary<string, string> _hyphenationDictionaryFiles;

        public CustomHyphenationDictionaryRegister(string dictionariesFolder)
        {
            _hyphenationDictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(dictionariesFolder, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(dictionariesFolder, "hyph_de_CH.dic") }
                // Add more language mappings here as needed.
            };
        }

        // Called by Aspose.Words when a dictionary for a language is required.
        public void RequestDictionary(string language)
        {
            Console.Write($"Hyphenation dictionary requested: {language}");

            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine(", already registered.");
                return;
            }

            // Register the dictionary if we know its location.
            if (_hyphenationDictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary available – optionally register a null dictionary to suppress further callbacks.
            Console.WriteLine(", no dictionary file known.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder that contains the source document and hyphenation dictionary files.
            string dataDir = @"C:\MyData";
            // Folder where the resulting document will be saved.
            string outputDir = @"C:\MyOutput";

            // Optional: pre‑register a dictionary that we know will be used.
            // This demonstrates registration by stream, but you can also use the file overload.
            using (Stream enUsStream = new FileStream(Path.Combine(dataDir, "hyph_en_US.dic"), FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", enUsStream);
            }

            // Set the callback that will lazily load dictionaries for other languages (e.g., German).
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

            // Load the document that contains text in a language for which we have not pre‑registered a dictionary.
            Document doc = new Document(Path.Combine(dataDir, "German text.docx"));

            // Enable automatic hyphenation for the whole document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Save the document – during layout the callback will be invoked for "de-CH".
            doc.Save(Path.Combine(outputDir, "HyphenatedOutput.pdf"));
        }
    }
}
