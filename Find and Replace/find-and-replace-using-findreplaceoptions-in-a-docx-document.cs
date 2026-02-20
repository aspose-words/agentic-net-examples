using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceDemo
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create FindReplaceOptions and configure desired behavior.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Perform a case‑sensitive search.
            MatchCase = true,
            // Replace only whole words.
            FindWholeWordsOnly = true,
            // Ignore text inside footnotes during the search.
            IgnoreFootnotes = true,
            // Use the forward direction (default) – can be changed to Backward if needed.
            Direction = FindReplaceDirection.Forward,
            // Specify that the replacement text is plain text.
            ReplacementFormat = ReplacementFormat.Text
        };

        // Define the text to find (as a regular expression) and the replacement string.
        Regex findPattern = new Regex(@"\bOldValue\b");
        string replacement = "NewValue";

        // Execute the find‑and‑replace operation with the configured options.
        doc.Range.Replace(findPattern, replacement, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
