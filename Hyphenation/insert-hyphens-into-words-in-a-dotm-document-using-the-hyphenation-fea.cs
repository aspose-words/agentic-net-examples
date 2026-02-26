using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDotm
{
    static void Main()
    {
        // Path to the folder that contains the input DOTM template and where the output will be saved.
        string dataDir = @"C:\Docs\";

        // Load the existing DOTM document.
        // This uses the Document(string) constructor – the provided load rule.
        Document doc = new Document(Path.Combine(dataDir, "Template.dotm"));

        // Enable automatic hyphenation for the whole document.
        // HyphenationOptions is accessed via the Document.HyphenationOptions property – the provided feature rule.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // maximum consecutive hyphenated lines
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin
        doc.HyphenationOptions.HyphenateCaps = true;       // hyphenate all‑caps words

        // Re‑layout the document so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the hyphenated document.
        // This uses the Document.Save(string) method – the provided save rule.
        doc.Save(Path.Combine(dataDir, "Hyphenated.docx"));
    }
}
