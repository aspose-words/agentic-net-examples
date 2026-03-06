using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text to demonstrate hyphenation.
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: fine‑tune hyphenation behavior.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch from right margin.
        doc.HyphenationOptions.HyphenateCaps = true;      // Hyphenate all‑caps words.

        // Save the resulting DOCX file.
        doc.Save("HyphenatedDocument.docx");
    }
}
