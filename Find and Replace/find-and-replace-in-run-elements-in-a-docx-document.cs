using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Simple string replace (case‑insensitive by default).
        // This will replace every occurrence of "_FullName_" with "John Doe"
        // even if the placeholder spans multiple Run nodes.
        int replacedCount = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Replacements made: {replacedCount}");

        // Example with FindReplaceOptions:
        // Replace the word "Numbers" with "Numbers" followed by a paragraph break.
        // The meta‑character "&p" inserts a paragraph break.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // ignore case
            FindWholeWordsOnly = true   // replace whole words only
        };
        doc.Range.Replace("Numbers", "Numbers&p", options);

        // Example using a regular expression to replace all digits with a line break.
        // The meta‑character "&l" inserts a manual line break.
        Regex digitPattern = new Regex(@"\d+");
        doc.Range.Replace(digitPattern, "&l");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
