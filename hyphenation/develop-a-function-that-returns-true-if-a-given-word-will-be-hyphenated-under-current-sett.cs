using System;
using System.IO;
using Aspose.Words;
using static Aspose.Words.Hyphenation; // Import static members of the Hyphenation class

public class Program
{
    // Path to the local hyphenation dictionary.
    private const string DictionaryPath = "hyph_en_US.dic";

    // Language code used for the dictionary.
    private const string Language = "en-US";

    // Ensures that a minimal dictionary file exists and is registered.
    private static void EnsureDictionaryRegistered()
    {
        if (!File.Exists(DictionaryPath))
        {
            // Create a deterministic dictionary with a few sample entries.
            // The format is: UTF-8 on the first line, then word=hyphenated‑pattern on subsequent lines.
            File.WriteAllText(DictionaryPath,
                "UTF-8\n" +
                "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
                "internationalization=in-ter-na-tion-al-i-za-tion\n" +
                "communication=com-mu-ni-ca-tion\n");
        }

        // Register the dictionary if it has not been registered yet.
        if (!IsDictionaryRegistered(Language))
            RegisterDictionary(Language, DictionaryPath);
    }

    // Returns true if the supplied word has a hyphenation entry in the registered dictionary.
    public static bool WillHyphenate(string word)
    {
        if (string.IsNullOrWhiteSpace(word))
            return false;

        EnsureDictionaryRegistered();

        // Simple lookup: the dictionary file contains lines "word=pattern".
        // If the word appears before the '=', we consider it hyphenatable.
        foreach (var line in File.ReadAllLines(DictionaryPath))
        {
            // Skip the first line which contains the encoding marker.
            if (line.StartsWith("UTF-8", StringComparison.OrdinalIgnoreCase))
                continue;

            var trimmed = line.Trim();
            if (trimmed.Length == 0)
                continue;

            var parts = trimmed.Split('=', 2);
            if (parts.Length == 2 && string.Equals(parts[0], word, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    // Demonstrates the usage of WillHyphenate.
    public static void Main()
    {
        // Create a blank document and enable automatic hyphenation.
        var doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to force hyphenation when possible.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add a sample paragraph containing words that may be hyphenated.
        var builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document to trigger layout (the file is not required for the core logic).
        doc.Save("HyphenationDemo.pdf");

        // Test the helper function with various words.
        string[] testWords = { "communication", "extraordinarycharacteristically", "unknownword" };
        foreach (var w in testWords)
        {
            bool canHyphenate = WillHyphenate(w);
            Console.WriteLine($"Word \"{w}\" hyphenatable: {canHyphenate}");
        }

        // Validate that the PDF was created.
        if (!File.Exists("HyphenationDemo.pdf"))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
