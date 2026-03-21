using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;

class HyphenationExample
{
    // Language code used by Aspose.Words for Spanish (Spain).
    private const string SpanishLanguageCode = "es-ES";

    static void Main()
    {
        // 1. Create a simple dictionary file for Spanish.
        string tempDictionaryPath = CreateSimpleSpanishDictionary();

        // 2. Register the dictionary for Spanish.
        Hyphenation.RegisterDictionary(SpanishLanguageCode, tempDictionaryPath);

        // Optional: set a callback to handle future requests for other languages.
        Hyphenation.Callback = new AutoHyphenationRegister();

        // 3. Create a document with Spanish text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font locale to Spanish so that hyphenation uses the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;
        builder.Writeln("Este es un ejemplo de texto en español que será dividido en sílabas mediante la hyphenación.");

        // 4. Save the document (PDF shows hyphenation on layout).
        doc.Save("SpanishHyphenated.pdf");

        // Clean up the temporary dictionary file.
        File.Delete(tempDictionaryPath);
    }

    // Creates a minimal Spanish hyphenation dictionary file and returns its path.
    private static string CreateSimpleSpanishDictionary()
    {
        // Very small dictionary content – just enough for the example text.
        // The first line is the number of patterns (can be 0 for a minimal file).
        string[] lines =
        {
            "0"
        };

        string tempPath = Path.GetTempFileName();
        File.WriteAllLines(tempPath, lines);
        return tempPath;
    }

    // Simple callback that registers dictionaries on demand if they are not already registered.
    private class AutoHyphenationRegister : IHyphenationCallback
    {
        // Mapping of language codes to local dictionary file paths.
        private readonly Dictionary<string, string> _dictionaryFiles = new Dictionary<string, string>
        {
            // For this example we only have Spanish; other languages will be ignored.
            { SpanishLanguageCode, CreateSimpleSpanishDictionary() }
        };

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a known file for the requested language.
            if (_dictionaryFiles.TryGetValue(language, out string filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
            }
            else
            {
                // Register a null dictionary to suppress further callbacks for this language.
                Hyphenation.RegisterDictionary(language, (string)null);
            }
        }
    }
}
