using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Max 2 consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch from right margin.
        doc.HyphenationOptions.HyphenateCaps = true;      // Hyphenate all‑caps words.

        // Optional: register a hyphenation dictionary for a specific language (e.g., German).
        // This step is only needed if the document's language is not covered by the built‑in dictionaries.
        // string dicPath = Path.Combine(Environment.CurrentDirectory, "hyph_de_CH.dic");
        // Hyphenation.RegisterDictionary("de-CH", dicPath);

        // Save the document to a DOCX file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HyphenatedDocument.docx");
        doc.Save(outputPath);
    }
}
