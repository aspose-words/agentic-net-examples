using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Optional: capture warnings that may occur during dictionary registration.
        var warningInfo = new WarningInfoCollection();
        Hyphenation.WarningCallback = warningInfo;

        // Register an English dictionary upfront (example).
        using (Stream enStream = new FileStream(Path.Combine(MyDir, "hyph_en_US.dic"), FileMode.Open))
        {
            Hyphenation.RegisterDictionary("en-US", enStream);
        }

        // Set the callback that will be invoked when a required dictionary is missing.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

        // Load a document that contains text in a language for which we have a dictionary (e.g., German).
        Document doc = new Document(Path.Combine(MyDir, "German text.docx"));

        // Enable automatic hyphenation so that the registered dictionaries are applied.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document; the callback will register the needed dictionary during layout.
        doc.Save(Path.Combine(ArtifactsDir, "HyphenatedOutput.pdf"));
    }

    // Adjust these paths to point to your data and output folders.
    private static readonly string MyDir = @"C:\Data";
    private static readonly string ArtifactsDir = @"C:\Output";

    // Implementation of the hyphenation callback.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister()
        {
            // Map language codes to local dictionary file paths.
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(MyDir, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(MyDir, "hyph_de_CH.dic") }
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
            }
            else
            {
                // No dictionary available for this language.
                Console.WriteLine(", no dictionary file known.");
            }
        }
    }
}
