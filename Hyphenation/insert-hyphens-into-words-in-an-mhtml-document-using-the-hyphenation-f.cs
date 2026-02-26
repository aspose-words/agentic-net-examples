using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Input and output MHTML file paths.
        string inputPath = @"C:\Docs\input.mhtml";
        string outputPath = @"C:\Docs\output.mhtml";

        // Load the MHTML document.
        Document doc = new Document(inputPath);

        // Ensure a hyphenation dictionary for the document language is registered.
        // Here we use English (US) as an example.
        const string language = "en-US";
        if (!Hyphenation.IsDictionaryRegistered(language))
        {
            // Path to the hyphenation dictionary file (OpenOffice format).
            string dictPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "hyph_en_US.dic");
            Hyphenation.RegisterDictionary(language, dictPath);
        }

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional settings: how far from the right margin hyphenation is allowed,
        // maximum consecutive hyphenated lines, and whether to hyphenate all‑caps words.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch (720 / 20 points)
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // limit consecutive hyphens
        doc.HyphenationOptions.HyphenateCaps = true;         // hyphenate capitalized words

        // Save the modified document back to MHTML format.
        doc.Save(outputPath);
    }
}
