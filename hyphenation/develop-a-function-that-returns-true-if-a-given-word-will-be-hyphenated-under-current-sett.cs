using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    // Returns true if the current document settings allow hyphenation for the specified language.
    // The check is based on whether automatic hyphenation is enabled and a hyphenation
    // dictionary for the language is registered.
    public static bool WillHyphenate(string word, Document doc, string languageCode)
    {
        // Auto hyphenation must be turned on.
        if (!doc.HyphenationOptions.AutoHyphenation)
            return false;

        // A hyphenation dictionary for the language must be registered.
        return Hyphenation.IsDictionaryRegistered(languageCode);
    }

    public static void Main()
    {
        // Prepare a folder for temporary artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a minimal hyphenation dictionary file for English (US).
        // The content does not need to contain real patterns for this demonstration.
        string dictPath = Path.Combine(artifactsDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath, "%%\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new document and enable automatic hyphenation.
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set the language of the text that will be added.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("Hyphenation");

        // Example word to test.
        string testWord = "Hyphenation";

        // Determine whether the word will be hyphenated under the current settings.
        bool canHyphenate = WillHyphenate(testWord, doc, "en-US");

        // Output the result.
        Console.WriteLine($"Will the word \"{testWord}\" be hyphenated? {canHyphenate}");

        // Save the document to demonstrate that the example runs without external input.
        string outputPath = Path.Combine(artifactsDir, "HyphenationDemo.docx");
        doc.Save(outputPath);
    }
}
