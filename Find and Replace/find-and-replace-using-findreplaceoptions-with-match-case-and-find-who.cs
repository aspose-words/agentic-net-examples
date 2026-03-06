using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        // Replace "input.docx" with the path to your source document.
        Document doc = new Document("input.docx");

        // Set up find-and-replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Enable case‑sensitive search.
            MatchCase = true,
            // Replace only whole words, not substrings.
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        // Replace "OldText" with "NewText". Adjust the strings as needed.
        int replacements = doc.Range.Replace("OldText", "NewText", options);

        // Optionally, output the number of replacements made.
        Console.WriteLine($"Replacements performed: {replacements}");

        // Save the modified document.
        // Replace "output.docx" with the desired output path.
        doc.Save("output.docx");
    }
}
