using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateRtf
{
    static void Main()
    {
        // Load the source RTF document.
        Document doc = new Document("Input.rtf");

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // limit consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin
        doc.HyphenationOptions.HyphenateCaps = true;       // hyphenate all‑caps words

        // Register an English (US) hyphenation dictionary if it is not already available.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            // Path to the OpenOffice‑format dictionary file.
            Hyphenation.RegisterDictionary("en-US", "hyph_en_US.dic");
        }

        // Save the modified document back to RTF – hyphens will be inserted where needed.
        doc.Save("Output.rtf");
    }
}
