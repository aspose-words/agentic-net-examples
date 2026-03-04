using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateMarkdown
{
    static void Main()
    {
        // Path to the source Markdown file.
        string inputPath = @"C:\Docs\sample.md";

        // Path to the resulting document (DOCX preserves hyphenation visually).
        string outputPath = @"C:\Docs\sample_hyphenated.docx";

        // Load the Markdown document. Aspose.Words detects the format from the file extension.
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Limit consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate all‑caps words.

        // Save the document. The hyphenation marks will be applied during layout.
        doc.Save(outputPath);
    }
}
