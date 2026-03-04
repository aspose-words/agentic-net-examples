using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;

class HyphenationExample
{
    // Paths to your data folders – adjust as needed.
    private const string MyDir = @"C:\Data\";
    private const string ArtifactsDir = @"C:\Output\";

    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Build the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Add a paragraph with German text.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("Dies ist ein Beispieltext, der die Silbentrennung demonstriert.");

        // 3. Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // 4. Register a custom hyphenation dictionary for German (Switzerland).
        // The dictionary file must be in OpenOffice format.
        string germanDictionaryPath = Path.Combine(MyDir, "hyph_de_CH.dic");
        Hyphenation.RegisterDictionary("de-CH", germanDictionaryPath);

        // 5. Optional: set a callback to load dictionaries on demand for other languages.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

        // 6. Save the document as DOCX (the create‑load‑save lifecycle is respected).
        doc.Save(Path.Combine(ArtifactsDir, "HyphenatedDocument.docx"));
    }

    // Callback implementation that registers dictionaries when the layout engine requests them.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles = new Dictionary<string, string>
        {
            { "en-US", Path.Combine(MyDir, "hyph_en_US.dic") },
            { "de-CH", Path.Combine(MyDir, "hyph_de_CH.dic") }
            // Add more language‑code / file‑path pairs as needed.
        };

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a known file for the requested language.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                return;
            }

            // If no dictionary is available, register a null dictionary to suppress further callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
