using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationStatusExample
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent = "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ninternationalization=in-ter-na-tion-al-i-za-tion\ncommunication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document and add a paragraph with sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document (optional, just to demonstrate the lifecycle).
        const string outputPath = "HyphenationStatus.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Retrieve the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        string paragraphText = paragraph.GetText().Trim();

        // Split the paragraph into words.
        string[] words = paragraphText.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        // Determine the language code for hyphenation (using the document's locale if set, otherwise default to en-US).
        string language = "en-US";

        // Log hyphenation status for each word.
        Console.WriteLine($"Hyphenation dictionary registered for '{language}': {Hyphenation.IsDictionaryRegistered(language)}");
        Console.WriteLine($"Automatic hyphenation enabled: {doc.HyphenationOptions.AutoHyphenation}");
        Console.WriteLine("Word hyphenation status:");
        foreach (string word in words)
        {
            // For demonstration, we consider a word hyphenatable if the dictionary is registered.
            bool hyphenatable = Hyphenation.IsDictionaryRegistered(language);
            Console.WriteLine($"- \"{word}\": {(hyphenatable ? "Hyphenation possible" : "Hyphenation not available")}");
        }
    }
}
