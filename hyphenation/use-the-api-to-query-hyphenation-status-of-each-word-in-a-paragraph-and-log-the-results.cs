using System;
using System.Globalization;
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

        // Configure the document to use automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Narrow the page width to force line wrapping and possible hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points (~2.78 inches)

        // Set the language of the paragraph to English (US) and write sample text.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln(
            "Hyphenation demonstration with extraordinarilylongwordthatmightneedhyphenation " +
            "and anotherverylongwordthatcouldbehyphenated to illustrate the process.");

        // Path for the output document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationStatus.docx");
        doc.Save(outputPath);

        // Verify that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // Determine if a hyphenation dictionary for the paragraph's language is registered.
        string language = "en-US";
        bool dictionaryRegistered = Hyphenation.IsDictionaryRegistered(language);

        // Log the hyphenation status for each word in the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        string paragraphText = paragraph.GetText().TrimEnd('\r', '\a'); // Remove trailing control characters.

        // Split the paragraph into words using common delimiters.
        char[] delimiters = new char[] { ' ', '\t', '\n', '\r', ',', '.', ';', ':', '!' , '?' };
        string[] words = paragraphText.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        Console.WriteLine($"Hyphenation dictionary for \"{language}\" registered: {dictionaryRegistered}");
        Console.WriteLine("Word hyphenation status (true = hyphenation could be applied):");
        foreach (string word in words)
        {
            // Hyphenation can be applied if automatic hyphenation is enabled and a dictionary is registered.
            bool hyphenationPossible = doc.HyphenationOptions.AutoHyphenation && dictionaryRegistered;
            Console.WriteLine($"  {word}: {hyphenationPossible}");
        }

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
