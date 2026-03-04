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

        // Add a long paragraph that will benefit from hyphenation.
        builder.Font.Size = 24;
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Enable automatic hyphenation and configure its behavior.
        doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;      // Limit consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;           // Hyphenate all‑caps words.

        // Save the document to a DOCX file.
        doc.Save("HyphenatedDocument.docx");
    }
}
