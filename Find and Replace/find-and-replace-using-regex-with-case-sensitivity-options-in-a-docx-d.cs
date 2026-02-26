using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithRegex
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Define the regular expression pattern to search for.
        // Example: find all occurrences of the word "color" with any case.
        Regex regexPattern = new Regex(@"\bcolor\b", RegexOptions.Compiled);

        // Replacement text.
        string replacement = "colour";

        // ---------- Case‑sensitive replace ----------
        FindReplaceOptions caseSensitiveOptions = new FindReplaceOptions
        {
            // Enable case‑sensitive matching.
            MatchCase = true
        };

        // Perform the replace operation with case sensitivity.
        int replacedCaseSensitive = doc.Range.Replace(regexPattern, replacement, caseSensitiveOptions);
        Console.WriteLine($"Case‑sensitive replacements made: {replacedCaseSensitive}");

        // ---------- Case‑insensitive replace ----------
        FindReplaceOptions caseInsensitiveOptions = new FindReplaceOptions
        {
            // Disable case‑sensitive matching (default is false, but set explicitly for clarity).
            MatchCase = false
        };

        // Perform the replace operation without case sensitivity.
        int replacedCaseInsensitive = doc.Range.Replace(regexPattern, replacement, caseInsensitiveOptions);
        Console.WriteLine($"Case‑insensitive replacements made: {replacedCaseInsensitive}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
