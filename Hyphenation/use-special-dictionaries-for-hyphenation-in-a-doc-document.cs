using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Directory that contains the source document and hyphenation dictionary files.
        string dataDir = @"C:\Data\";

        // Input Word document that contains text in the language we want to hyphenate.
        string inputDocPath = Path.Combine(dataDir, "German text.docx");

        // Output PDF file where hyphenated text will be rendered.
        string outputPdfPath = Path.Combine(dataDir, "HyphenatedOutput.pdf");

        // Register a hyphenation dictionary for German (Switzerland) language.
        // This dictionary will be used automatically when the document is laid out.
        Hyphenation.RegisterDictionary("de-CH", Path.Combine(dataDir, "hyph_de_CH.dic"));

        // Set a callback to handle any other language requests during layout.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dataDir);

        // Load the source document.
        Document doc = new Document(inputDocPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF; hyphenation will be applied during the save operation.
        doc.Save(outputPdfPath);
    }

    // Callback implementation that registers dictionaries on demand.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string basePath)
        {
            // Map language codes to dictionary file locations.
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(basePath, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(basePath, "hyph_de_CH.dic") }
                // Add more language mappings here if needed.
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

            // Register the dictionary if we have a known file for the requested language.
            if (_dictionaryFiles.TryGetValue(language, out string filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
            }
            else
            {
                // No dictionary available; further requests for this language will be ignored.
                Console.WriteLine(", no dictionary found.");
            }
        }
    }
}
