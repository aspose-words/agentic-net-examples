using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        string inputPath = @"C:\Docs\Input.docx";
        Document doc = new Document(inputPath);

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: fine‑tune hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate all‑caps words.

        // Save the document; hyphenation will be applied during layout.
        string outputPath = @"C:\Docs\Output_Hyphenated.docx";
        doc.Save(outputPath);
    }
}
