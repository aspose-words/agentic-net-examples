using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationStatusExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with words that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary and enable automatic hyphenation.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document.
        const string outputPath = "HyphenatedDocument.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Retrieve the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        string paragraphText = paragraph.GetText().Trim();

        // Split the paragraph into individual words.
        string[] words = paragraphText.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        // Log hyphenation status for each word.
        Console.WriteLine("Hyphenation status per word (dictionary registered for en-US):");
        foreach (string word in words)
        {
            // The Hyphenation API does not expose per‑word status, so we report whether a dictionary is available.
            bool dictionaryRegistered = Hyphenation.IsDictionaryRegistered("en-US");
            Console.WriteLine($"Word: \"{word}\", HyphenationPossible: {dictionaryRegistered}");
        }
    }
}
