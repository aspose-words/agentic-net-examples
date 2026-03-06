using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;          // Hyphenate words in all caps.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max consecutive hyphenated lines.

        // Ensure a hyphenation dictionary for the document language is registered.
        // Here we use English (US) as an example.
        const string language = "en-US";
        if (!Hyphenation.IsDictionaryRegistered(language))
        {
            // Register the dictionary file (OpenOffice format) located alongside the executable.
            Hyphenation.RegisterDictionary(language, "hyph_en_US.dic");
        }

        // Rebuild the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the document with hyphens inserted.
        doc.Save("HyphenatedDocument.dotx");
    }
}
