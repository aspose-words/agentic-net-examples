using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Simple string replace (case‑insensitive by default).
        // Replaces every occurrence of the placeholder _FullName_ with "John Doe".
        int simpleReplacements = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Simple replacements made: {simpleReplacements}");

        // Replace with additional options:
        // - MatchCase = true makes the search case‑sensitive.
        // - FindWholeWordsOnly = true ensures only whole word matches are replaced.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };
        int optionReplacements = doc.Range.Replace("Aspose", "Aspose.Words", options);
        Console.WriteLine($"Option‑based replacements made: {optionReplacements}");

        // Regular‑expression replace:
        // Replace every sequence of digits with a paragraph break (&p).
        int regexReplacements = doc.Range.Replace(new Regex(@"\d+"), "&p", new FindReplaceOptions());
        Console.WriteLine($"Regex replacements made: {regexReplacements}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
