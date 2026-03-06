using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        string inputPath = @"C:\Docs\input.doc"; // <-- replace with your DOC file path
        Document doc = new Document(inputPath);

        // Turn on automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Register the dictionaries that will be used during layout.
        // The streams are opened with read‑only access and will be closed by the layout engine.
        string enUsDictionaryPath = @"C:\Hyphenation\en-US.dic"; // <-- replace with your OpenOffice dictionary
        string deChDictionaryPath = @"C:\Hyphenation\de-CH.dic"; // <-- replace with your OpenOffice dictionary

        Hyphenation.RegisterDictionary("en-US", File.OpenRead(enUsDictionaryPath));
        Hyphenation.RegisterDictionary("de-CH", File.OpenRead(deChDictionaryPath));

        // Optional: set a callback to load dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

        // Save the hyphenated document.
        string outputPath = @"C:\Docs\output.docx"; // <-- replace with desired output path
        doc.Save(outputPath);
    }

    // Callback implementation that registers dictionaries when requested by the layout engine.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles = new Dictionary<string, string>
        {
            { "en-US", @"C:\Hyphenation\en-US.dic" },
            { "de-CH", @"C:\Hyphenation\de-CH.dic" }
        };

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, no action is needed.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a known file for the requested language.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, File.OpenRead(filePath));
            }
            // If no dictionary is known for the language, the layout will proceed without hyphenation for that language.
        }
    }
}
