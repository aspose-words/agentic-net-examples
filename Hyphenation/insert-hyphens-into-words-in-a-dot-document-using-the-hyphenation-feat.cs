using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Folder where the output file will be saved.
        string artifactsDir = @"C:\Artifacts\";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long paragraph that will require hyphenation.
        builder.Font.Size = 24;
        builder.Writeln(
            "Hyphenation demonstrates how longwordslikecharacteristicscanbeautomaticallyhyphenatedwhenthelineexceedsmargin.");

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;     // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;          // Hyphenate all‑caps words.

        // Save the document as a DOT template.
        string outputPath = Path.Combine(artifactsDir, "HyphenatedTemplate.dot");
        doc.Save(outputPath);
    }
}
