using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateMhtml
{
    static void Main()
    {
        // Load the source MHTML document.
        Document doc = new Document("Input.mhtml");

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // limit consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin
        doc.HyphenationOptions.HyphenateCaps = true;       // hyphenate all‑caps words

        // Register an English hyphenation dictionary if it is not already available.
        // This step is required when the document language does not have a built‑in dictionary.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            using (FileStream stream = new FileStream("hyph_en_US.dic", FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", stream);
            }
        }

        // Re‑layout the document so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the hyphenated document. The format is inferred from the file extension.
        doc.Save("Output.docx");
    }
}
