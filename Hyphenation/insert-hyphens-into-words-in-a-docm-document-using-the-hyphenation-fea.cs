using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDocm
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = "input.docm";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // limit consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch from the right margin
        doc.HyphenationOptions.HyphenateCaps = true;      // hyphenate all‑caps words

        // Save the modified document as a DOCM file.
        string outputPath = "output.docm";
        doc.Save(outputPath);
    }
}
