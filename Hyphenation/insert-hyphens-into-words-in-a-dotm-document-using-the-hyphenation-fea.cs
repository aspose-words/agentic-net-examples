using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDotm
{
    static void Main()
    {
        // Load the existing DOTM (macro-enabled template) document.
        Document doc = new Document("Template.dotm");

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: fine‑tune hyphenation behavior.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;       // Hyphenate all‑caps words.

        // Save the document with hyphens applied.
        doc.Save("HyphenatedTemplate.dotm");
    }
}
