using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph containing short and long words.
        // Short words (e.g., "A", "short") will not be hyphenated,
        // while longer words may be split according to the hyphenation dictionary.
        builder.Font.Size = 24;
        builder.Writeln("A short example with hyphenation: extraordinary, international, supercalifragilisticexpialidocious.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Aspose.Words does not provide a public property to set the minimum word length
        // for hyphenation directly. The internal algorithm typically skips words shorter
        // than five characters, so short words in the paragraph above will remain un‑hyphenated.
        // If a custom hyphenation dictionary is used, its patterns can influence this behavior.

        // Save the resulting document.
        doc.Save("HyphenationMinimumWordLength.docx");
    }
}
