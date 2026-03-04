using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class SmartTagFindReplace
{
    static void Main()
    {
        // Load the DOCX document that contains smart tags.
        Document doc = new Document("InputWithSmartTags.docx");

        // Create FindReplaceOptions.
        // Setting IgnoreStructuredDocumentTags to true treats the content of smart tags as plain text,
        // allowing the replace operation to work inside them.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreStructuredDocumentTags = true,
            MatchCase = false,               // case‑insensitive replace
            FindWholeWordsOnly = false       // replace even if the pattern is part of a larger word
        };

        // Define the text to find and its replacement.
        string pattern = "_FullName_";      // example placeholder inside a smart tag
        string replacement = "John Doe";

        // Perform the find‑and‑replace on the whole document range.
        int replacedCount = doc.Range.Replace(pattern, replacement, options);

        Console.WriteLine($"Replacements made: {replacedCount}");

        // Save the modified document.
        doc.Save("OutputWithSmartTagsReplaced.docx");
    }
}
