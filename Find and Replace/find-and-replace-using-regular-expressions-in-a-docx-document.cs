using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class RegexFindReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define a regular expression pattern.
        // This example finds all occurrences of one or more digits.
        Regex regexPattern = new Regex(@"\d+");

        // Define the replacement string.
        // "&p" inserts a paragraph break for each match.
        string replacement = "&p";

        // Optional: configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions();
        // Example: make the operation case‑insensitive (default for regex).
        options.MatchCase = false;

        // Perform the find‑and‑replace operation on the whole document range.
        int replacementsMade = doc.Range.Replace(regexPattern, replacement, options);

        Console.WriteLine($"Number of replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
