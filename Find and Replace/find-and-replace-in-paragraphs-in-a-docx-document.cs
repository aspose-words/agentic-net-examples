using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Simple find‑and‑replace: replace the placeholder _FullName_ with an actual name.
        // Returns the number of replacements made.
        int simpleReplacements = doc.Range.Replace("_FullName_", "John Doe");

        // Find‑and‑replace with additional options.
        // This example replaces the word "important" with "crucial",
        // ignores case, replaces only whole words, and right‑aligns every paragraph
        // that contains a match.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,                 // Case‑insensitive search.
            FindWholeWordsOnly = true          // Replace only whole‑word matches.
        };
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;

        int advancedReplacements = doc.Range.Replace("important", "crucial", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
