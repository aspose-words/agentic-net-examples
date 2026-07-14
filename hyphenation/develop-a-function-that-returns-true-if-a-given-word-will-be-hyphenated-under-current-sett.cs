using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationDemo
{
    // Returns true if the specified word can be hyphenated with the current settings.
    public static bool WillHyphenate(string word, string language = "en-US")
    {
        // Create a temporary hyphenation dictionary file if one does not already exist for this language.
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), $"hyph_{language}.dic");
        if (!File.Exists(dictPath))
        {
            // Simple dictionary format: first line is the encoding, subsequent lines are word=hyphenated-pattern.
            string hyphenatedPattern = string.Join("-", word.ToCharArray());
            string content = $"UTF-8{Environment.NewLine}{word}={hyphenatedPattern}{Environment.NewLine}";
            File.WriteAllText(dictPath, content);
        }

        // Register the dictionary for the requested language.
        if (!Hyphenation.IsDictionaryRegistered(language))
            Hyphenation.RegisterDictionary(language, dictPath);

        // Create a blank document and enable automatic hyphenation.
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write the test word using the appropriate locale.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo(language).LCID;
        builder.Writeln(word);

        // Determine if hyphenation can occur: a dictionary must be registered and auto‑hyphenation enabled.
        bool canHyphenate = Hyphenation.IsDictionaryRegistered(language) && doc.HyphenationOptions.AutoHyphenation;

        // Clean up the temporary dictionary file.
        try { File.Delete(dictPath); } catch { /* ignore cleanup errors */ }

        return canHyphenate;
    }

    public static void Main()
    {
        string testWord = "extraordinarycharacteristically";
        bool result = WillHyphenate(testWord, "en-US");
        Console.WriteLine($"Will the word \"{testWord}\" be hyphenated? {result}");
    }
}
