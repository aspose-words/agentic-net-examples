using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the source MHTML document.
        Document doc = new Document("Input.mhtml");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // limit consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin
        doc.HyphenationOptions.HyphenateCaps = true;       // hyphenate all‑caps words

        // Register an English hyphenation dictionary if it is not already available.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            using (FileStream stream = new FileStream("hyph_en_US.dic", FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", stream);
            }
        }

        // Rebuild the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the document back to MHTML format with hyphens inserted.
        doc.Save("Output.mhtml");
    }
}
