using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the existing DOCM file.
        Document doc = new Document("Input.docm");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate all‑caps words.

        // Rebuild the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the modified document back to DOCM format.
        doc.Save("Output.docm");
    }
}
