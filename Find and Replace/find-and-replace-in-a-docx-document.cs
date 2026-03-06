using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Simple string replace (case‑insensitive by default).
        int simpleCount = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Simple replacements made: {simpleCount}");

        // Replace with additional options: case‑sensitive and whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };
        int caseSensitiveCount = doc.Range.Replace("Ruby", "Jade", options);
        Console.WriteLine($"Case‑sensitive replacements made: {caseSensitiveCount}");

        // Replace using a regular expression; replace each number with a paragraph break.
        int regexCount = doc.Range.Replace(new Regex(@"\d+"), "&p");
        Console.WriteLine($"Regex replacements made: {regexCount}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
