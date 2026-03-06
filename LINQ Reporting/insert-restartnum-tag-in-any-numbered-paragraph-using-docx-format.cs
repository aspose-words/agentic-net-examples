using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertRestartNumTag
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define replace options (case‑insensitive, not whole‑word only).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Regex that matches any paragraph that starts with a number followed by a period
        // (e.g., "1.", "  23.", etc.). The (?m) flag enables multiline mode so ^ matches
        // the start of each line/paragraph.
        Regex numberedParagraphPattern = new Regex(@"(?m)^\s*\d+\.", RegexOptions.Multiline);

        // Insert the <restartNum> tag before the matched numbering.
        // The replacement string uses $0 to keep the original matched text.
        int replacements = doc.Range.Replace(numberedParagraphPattern, "<restartNum>$0", options);

        Console.WriteLine($"Number of paragraphs updated: {replacements}");

        // Save the modified document in DOCX format.
        doc.Save("Output.docx");
    }
}
