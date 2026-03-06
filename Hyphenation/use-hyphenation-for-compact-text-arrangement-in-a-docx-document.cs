using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph with long text that will benefit from hyphenation.
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;               // Turn on hyphenation.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;           // Allow at most two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;               // Set hyphenation zone to 0.5 inch (720/1440 points).
        doc.HyphenationOptions.HyphenateCaps = true;                // Hyphenate words written in all caps.

        // Save the document as DOCX.
        doc.Save("HyphenatedDocument.docx");
    }
}
