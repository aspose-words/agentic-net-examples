using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceDemo
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // --------------------------------------------------------------------
        // Example 1: Case‑insensitive replace.
        // --------------------------------------------------------------------
        // Regex pattern to find the word "example" (any case).
        Regex patternInsensitive = new Regex(@"example", RegexOptions.None);

        // Configure find/replace options: case‑insensitive (MatchCase = false).
        FindReplaceOptions optionsInsensitive = new FindReplaceOptions
        {
            MatchCase = false   // false = ignore case.
        };

        // Replace all matches with the word "sample".
        doc.Range.Replace(patternInsensitive, "sample", optionsInsensitive);

        // --------------------------------------------------------------------
        // Example 2: Case‑sensitive replace.
        // --------------------------------------------------------------------
        // Regex pattern to find the exact case "Test".
        Regex patternSensitive = new Regex(@"Test", RegexOptions.None);

        // Configure find/replace options: case‑sensitive (MatchCase = true).
        FindReplaceOptions optionsSensitive = new FindReplaceOptions
        {
            MatchCase = true    // true = respect case.
        };

        // Replace all matches with the word "Exam".
        doc.Range.Replace(patternSensitive, "Exam", optionsSensitive);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
