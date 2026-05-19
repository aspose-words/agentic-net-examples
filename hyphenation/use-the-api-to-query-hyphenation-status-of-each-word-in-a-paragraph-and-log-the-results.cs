using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationStatusExample
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

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new document and add a paragraph with words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document.
        const string outputPath = "HyphenationStatus.docx";
        doc.Save(outputPath);

        // Verify that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved as expected.");

        // Query hyphenation status for each word in the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        foreach (Run run in paragraph.Runs)
        {
            string[] words = run.Text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string word in words)
            {
                // Hyphenation is possible if a dictionary for the language is registered.
                bool hyphenationAvailable = Hyphenation.IsDictionaryRegistered("en-US");
                Console.WriteLine($"Word: '{word}' – Hyphenation dictionary registered: {hyphenationAvailable}");
            }
        }
    }
}
