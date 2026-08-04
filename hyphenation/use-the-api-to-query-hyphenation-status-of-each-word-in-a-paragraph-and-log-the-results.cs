using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictionaryPath = "hyph_en_US.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate words for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary was registered successfully.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document and add a paragraph with words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale for the paragraph to English (US) and enable automatic hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write sample text.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document (required by the lifecycle rules).
        const string outputPath = "HyphenationStatus.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Retrieve the paragraph text (trim the trailing paragraph mark).
        string paragraphText = doc.FirstSection.Body.FirstParagraph.GetText().TrimEnd('\r', '\a');

        // Split the paragraph into individual words.
        string[] words = paragraphText.Split(
            new[] { ' ', '\t', '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);

        // Log the hyphenation status for each word.
        bool hyphenationEnabled = doc.HyphenationOptions.AutoHyphenation &&
                                  Hyphenation.IsDictionaryRegistered("en-US");

        foreach (string word in words)
        {
            Console.WriteLine($"Word: \"{word}\", HyphenationEnabled: {hyphenationEnabled}");
        }
    }
}
