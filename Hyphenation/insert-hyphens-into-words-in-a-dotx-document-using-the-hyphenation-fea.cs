using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDotx
{
    static void Main()
    {
        // Load the existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch from right margin.
        doc.HyphenationOptions.HyphenateCaps = true;      // Hyphenate all‑caps words.

        // Save the document with hyphenation applied.
        doc.Save("HyphenatedDocument.docx");
    }
}
