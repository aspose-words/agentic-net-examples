using System;
using Aspose.Words;
using Aspose.Words.Settings;
using System.Globalization;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;       // Hyphenate all‑caps words.

        // Force layout rebuild so hyphens are inserted.
        doc.UpdatePageLayout();

        // Save the document with hyphens applied.
        doc.Save("HyphenatedDocument.dotx");
    }
}
