using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words; // Hyphenation class resides here
using System.Globalization;

class HyphenateWordml
{
    static void Main()
    {
        // Load the WORDML (or DOCX) document.
        Document doc = new Document("Input.docx"); // replace with your WORDML file path

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // limit consecutive hyphenated lines
        doc.HyphenationOptions.HyphenateCaps = true;         // hyphenate all‑caps words

        // Ensure a hyphenation dictionary is registered for the document's language.
        // Example: register the English (US) dictionary if it is not already available.
        const string language = "en-US";
        if (!Hyphenation.IsDictionaryRegistered(language))
        {
            // Provide the path to the .dic file that contains hyphenation patterns.
            Hyphenation.RegisterDictionary(language, "hyph_en_US.dic");
        }

        // Save the document. Hyphenation will be applied during layout (e.g., when saving to PDF or DOCX).
        doc.Save("Output.docx"); // replace with desired output path
    }
}
