using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class RegexFindReplaceDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing different capitalizations.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("the QUICK brown FOX jumps over the LAZY dog.");

        // Define a regular expression pattern to find the word "quick" (case‑insensitive part).
        Regex pattern = new Regex(@"\bquick\b", RegexOptions.IgnoreCase);

        // Replace with "swift" using case‑insensitive matching.
        FindReplaceOptions caseInsensitiveOptions = new FindReplaceOptions
        {
            MatchCase = false   // Ignore case while searching.
        };
        doc.Range.Replace(pattern, "swift", caseInsensitiveOptions);

        // Define a regular expression pattern to find the word "fox" (case‑sensitive part).
        Regex caseSensitivePattern = new Regex(@"\bFox\b"); // No RegexOptions.IgnoreCase

        // Replace with "wolf" using case‑sensitive matching.
        FindReplaceOptions caseSensitiveOptions = new FindReplaceOptions
        {
            MatchCase = true    // Enforce case sensitivity.
        };
        doc.Range.Replace(caseSensitivePattern, "wolf", caseSensitiveOptions);

        // Save the modified document.
        doc.Save("RegexFindReplaceResult.docx");
    }
}
