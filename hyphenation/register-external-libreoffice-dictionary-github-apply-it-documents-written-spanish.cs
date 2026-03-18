using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Collections.Generic;
using Aspose.Words;

class HyphenationExample
{
    // URL of the LibreOffice Spanish hyphenation dictionary on GitHub.
    private const string DictionaryUrl = "https://raw.githubusercontent.com/LibreOffice/dictionaries/master/es_ES/hyph_es_ES.dic";

    // Language code used by Aspose.Words for Spanish (Spain).
    private const string SpanishLanguageCode = "es-ES";

    static void Main()
    {
        // 1. Download the dictionary file to a temporary location.
        string tempDictionaryPath = DownloadDictionary(DictionaryUrl);

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

    // Downloads the dictionary from the given URL and returns the local file path.
    private static string DownloadDictionary(string url)
    {
        using (HttpClient client = new HttpClient())
        using (Stream stream = client.GetStreamAsync(url).Result)
        {
            string tempPath = Path.GetTempFileName();
            using (FileStream file = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
            return tempPath;
        }
    }

    // Simple callback that registers dictionaries on demand if they are not already registered.
    private class AutoHyphenationRegister : IHyphenationCallback
    {
        // Mapping of language codes to local dictionary file paths.
        private readonly Dictionary<string, string> _dictionaryFiles = new Dictionary<string, string>
        {
            { HyphenationExample.SpanishLanguageCode, HyphenationExample.DownloadDictionary(HyphenationExample.DictionaryUrl) }
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
