using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationDemo
{
    // Returns true if the given word can be hyphenated with the current document settings.
    // The check is based on whether automatic hyphenation is enabled and a dictionary
    // for the document language is registered.
    public static bool WillHyphenate(string word, Document doc)
    {
        // Determine the language code to use. Here we use the locale of the first run
        // if it is set; otherwise we fall back to "en-US".
        string language = "en-US";

        if (doc.FirstSection?.Body?.FirstParagraph?.Runs?.Count > 0)
        {
            var firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];
            if (firstRun.Font?.LocaleId != 0)
                language = CultureInfo.GetCultureInfo(firstRun.Font.LocaleId).Name;
        }

        // Hyphenation is possible only when a dictionary for the language is registered
        // and automatic hyphenation is turned on for the document.
        bool dictionaryRegistered = Hyphenation.IsDictionaryRegistered(language);
        bool autoHyphenation = doc.HyphenationOptions.AutoHyphenation;

        return dictionaryRegistered && autoHyphenation;
    }

    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\nexample=ex-am-ple\nextraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("example extraordinarycharacteristically");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to PDF to force layout processing.
        const string outputPdf = "hyphenated.pdf";
        doc.Save(outputPdf);
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF output was not created.");

        // Test the helper method.
        string testWord = "example";
        bool canHyphenate = WillHyphenate(testWord, doc);
        Console.WriteLine($"Can the word \"{testWord}\" be hyphenated? {canHyphenate}");
    }
}
