using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a FindReplaceOptions object to customize the replace operation.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Example: make the search case‑insensitive.
            MatchCase = false,
            // Example: replace only whole words.
            FindWholeWordsOnly = true
        };

        // Perform the find‑and‑replace on the whole document range.
        // Replace every occurrence of the placeholder "_FullName_" with "John Doe".
        int replacements = doc.Range.Replace("_FullName_", "John Doe", options);

        // Output the number of replacements made (optional).
        Console.WriteLine($"Replacements performed: {replacements}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
