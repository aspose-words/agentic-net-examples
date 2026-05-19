using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    // Path to the locally created hyphenation dictionary.
    private static readonly string DictionaryPath = "hyph_en_US.dic";

    // Cache of words that have hyphenation entries in the dictionary.
    private static HashSet<string> _dictionaryWords;

    public static void Main()
    {
        // 1. Create a minimal hyphenation dictionary file.
        // The first line is the encoding header required by the format.
        // Each subsequent line contains a word and its hyphenation pattern separated by '='.
        File.WriteAllText(
            DictionaryPath,
            @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
");

        // 2. Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", DictionaryPath);

        // 3. Enable automatic hyphenation in a document (required for layout processing).
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.Save("dummy.docx"); // The document must be saved to trigger layout.

        // 4. Demonstrate the WillHyphenate function with several words.
        string[] testWords = {
            "extraordinarycharacteristically",
            "communication",
            "unregisteredword",
            "Internationalization"
        };

        foreach (string word in testWords)
        {
            bool canHyphenate = WillHyphenate(word, "en-US");
            Console.WriteLine($"Word \"{word}\" hyphenates: {canHyphenate}");
        }

        // Clean up the temporary files (optional).
        if (File.Exists("dummy.docx")) File.Delete("dummy.docx");
        if (File.Exists(DictionaryPath)) File.Delete(DictionaryPath);
    }

    /// <summary>
    /// Determines whether the specified word will be hyphenated under the current
    /// hyphenation settings for the given language.
    /// </summary>
    /// <param name="word">The word to test.</param>
    /// <param name="language">The language code (e.g., "en-US").</param>
    /// <returns>True if the word has a hyphenation entry and hyphenation is enabled; otherwise false.</returns>
    private static bool WillHyphenate(string word, string language)
    {
        // Hyphenation can only occur if a dictionary for the language is registered.
        if (!Hyphenation.IsDictionaryRegistered(language))
            return false;

        // Ensure the dictionary words are loaded into the cache.
        EnsureDictionaryWordsLoaded();

        // The dictionary stores words in lower‑case; perform a case‑insensitive check.
        return _dictionaryWords.Contains(word.ToLowerInvariant());
    }

    /// <summary>
    /// Loads the words from the dictionary file into the static cache.
    /// </summary>
    private static void EnsureDictionaryWordsLoaded()
    {
        if (_dictionaryWords != null)
            return; // Already loaded.

        _dictionaryWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // The dictionary file must exist; otherwise, hyphenation cannot be evaluated.
        if (!File.Exists(DictionaryPath))
            return;

        using (var reader = new StreamReader(DictionaryPath))
        {
            // Skip the first line (encoding header).
            if (!reader.EndOfStream) reader.ReadLine();

            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Each line is of the form "word=pattern".
                int separatorIndex = line.IndexOf('=');
                if (separatorIndex > 0)
                {
                    string dictWord = line.Substring(0, separatorIndex).Trim();
                    if (!string.IsNullOrEmpty(dictWord))
                        _dictionaryWords.Add(dictWord);
                }
            }
        }
    }
}
