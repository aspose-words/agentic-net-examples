using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateTxt
{
    static void Main()
    {
        // Path to the source TXT file.
        string inputPath = @"C:\Docs\input.txt";

        // Path where the hyphenated TXT will be saved.
        string outputPath = @"C:\Docs\output.txt";

        // Load the plain‑text document.
        Document doc = new Document(inputPath);

        // Turn on automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: fine‑tune hyphenation behaviour.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // max consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;       // distance from right margin (in 1/20 pt)
        doc.HyphenationOptions.HyphenateCaps = true;       // hyphenate all‑caps words

        // Save the document back as TXT – hyphens will be inserted where the layout permits.
        doc.Save(outputPath);
    }
}
