using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text that will be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Turn on automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch from right margin.
        doc.HyphenationOptions.HyphenateCaps = true;      // Hyphenate all‑caps words.

        // Save the hyphenated document.
        doc.Save("HyphenatedDocument.docx");
    }
}
