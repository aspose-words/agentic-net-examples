using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceExample
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Perform a case‑insensitive search.
            MatchCase = false,
            // Replace only whole‑word matches.
            FindWholeWordsOnly = true
        };

        // Replace all occurrences of the placeholder with the desired text.
        int replacementsMade = doc.Range.Replace("_FullName_", "John Doe", options);

        // Optionally, output the number of replacements performed.
        Console.WriteLine($"Replacements made: {replacementsMade}");

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
