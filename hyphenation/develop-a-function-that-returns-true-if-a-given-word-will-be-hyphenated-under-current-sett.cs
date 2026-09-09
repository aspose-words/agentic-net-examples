using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    // Mapping from language code to the full path of the registered dictionary file.
    private static readonly Dictionary<string, string> _registeredDictionaries = new();

    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string language = "en-US";
        const string dictFileName = "hyph_en_US.dic";

        // Dictionary format: first line is the encoding, subsequent lines are "word=hy-phen-ated".
        string dictContent =
            "UTF-8\n" +
            "hyphenation=hy-phen-ation\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n";

        // Write the dictionary file to the local folder.
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary(language, dictFileName);
        _registeredDictionaries[language] = Path.GetFullPath(dictFileName);

        // Example words to test.
        string[] words = { "hyphenation", "extraordinarycharacteristically", "unregisteredword" };

        foreach (string w in words)
        {
            bool canHyphenate = WillHyphenate(w, language);
            Console.WriteLine($"Word \"{w}\" hyphenated under current settings: {canHyphenate}");
        }

        // Clean up the temporary dictionary file.
        if (File.Exists(dictFileName))
            File.Delete(dictFileName);
    }

    /// <summary>
    /// Determines whether the specified word will be hyphenated under the current hyphenation settings.
    /// The method checks if a hyphenation dictionary is registered for the given language and
    /// whether the dictionary contains an entry for the word.
    /// </summary>
    /// <param name="word">The word to test.</param>
    /// <param name="language">The language code (e.g., "en-US").</param>
    /// <returns>True if the word has a hyphenation entry in the registered dictionary; otherwise false.</returns>
    private static bool WillHyphenate(string word, string language)
    {
        if (string.IsNullOrEmpty(word) || string.IsNullOrEmpty(language))
            return false;

        // Verify that a dictionary for the language is registered.
        if (!Hyphenation.IsDictionaryRegistered(language))
            return false;

        // Retrieve the path of the dictionary file that was registered.
        if (!_registeredDictionaries.TryGetValue(language, out string dictPath) || !File.Exists(dictPath))
            return false;

        // Read the dictionary file and look for an entry matching the word (case‑insensitive).
        // Dictionary lines after the first line have the format: originalWord=hy‑phen‑ated
        foreach (string line in File.ReadLines(dictPath))
        {
            // Skip the encoding header line.
            if (line.StartsWith("UTF-", StringComparison.OrdinalIgnoreCase))
                continue;

            // Ignore empty lines.
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Split the line into the original word and its hyphenated form.
            int separatorIndex = line.IndexOf('=');
            if (separatorIndex <= 0)
                continue; // malformed line

            string original = line.Substring(0, separatorIndex).Trim();
            // Compare the original word with the input word.
            if (string.Equals(original, word, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        // No matching entry found.
        return false;
    }
}
