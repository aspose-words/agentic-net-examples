using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Loading; // Added for LoadOptions and LoadFormat

class Program
{
    static void Main()
    {
        // Input markdown file and output document paths.
        string inputPath = "input.md";
        string outputPath = "output.docx";

        // Load the markdown file into an Aspose.Words Document.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Markdown };
        Document doc = new Document(inputPath, loadOptions);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure hyphenation behavior.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate all‑caps words.

        // Force layout calculation so hyphenation is applied before saving.
        doc.UpdatePageLayout();

        // Save the resulting document (hyphens will appear where Word would break lines).
        doc.Save(outputPath);
    }
}
