using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateDotm
{
    static void Main()
    {
        // Path to the source DOTM template.
        string inputPath = @"C:\Docs\Template.dotm";

        // Path where the hyphenated document will be saved.
        string outputPath = @"C:\Docs\Hyphenated.docx";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        doc.HyphenationOptions.HyphenateCaps = true;           // Hyphenate words in all caps.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;         // 0.5 inch from the right margin.

        // Rebuild layout so hyphens are applied.
        doc.UpdatePageLayout();

        // Save the result.
        doc.Save(outputPath);
    }
}
