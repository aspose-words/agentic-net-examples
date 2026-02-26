using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDocument
{
    static void Main()
    {
        // Load the existing WORDML document.
        Document doc = new Document("Input.docx");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: fine‑tune hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate all‑caps words.

        // Save the document with hyphens applied.
        doc.Save("Output.docx");
    }
}
