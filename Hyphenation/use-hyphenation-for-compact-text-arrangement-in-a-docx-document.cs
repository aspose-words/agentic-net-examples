using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a long paragraph to demonstrate hyphenation.
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;     // Limit consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;           // Hyphenate all‑caps words.

        // Register an English (US) hyphenation dictionary if it hasn't been registered yet.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            using (Stream dictStream = new FileStream("hyph_en_US.dic", FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", dictStream);
            }
        }

        // Save the document to a DOCX file.
        doc.Save("HyphenatedDocument.docx");
    }
}
