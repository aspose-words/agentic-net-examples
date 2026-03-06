using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenateMarkdown
{
    static void Main()
    {
        // Path to the input Markdown file.
        const string inputPath = @"C:\Docs\input.md";

        // Path to the output document (Word format will show hyphens when opened in Word).
        const string outputPath = @"C:\Docs\output.docx";

        // Load the Markdown document.
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: set hyphenation parameters.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // Max consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;       // Hyphenate all‑caps words.

        // Ensure the document language matches a registered hyphenation dictionary.
        // Here we use US English; Aspose.Words includes the built‑in dictionary for this locale.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = new CultureInfo("en-US").LCID;
        }

        // Save the document. The layout engine will insert hyphens where needed.
        doc.Save(outputPath);
    }
}
