using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceDemo
{
    static void Main()
    {
        // Load the DOCX document (lifecycle rule: load)
        Document doc = new Document("Input.docx");

        // Example 1: Simple string replace throughout the whole document (case‑insensitive)
        // Replaces all occurrences of the placeholder "_FullName_" with "John Doe".
        int count1 = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Simple replace made {count1} replacements.");

        // Example 2: Replace with additional options (match whole words only, case‑sensitive)
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,
            FindWholeWordsOnly = true
        };
        // Replace the word "Ruby" only when it appears as a whole word and with exact case.
        int count2 = doc.Range.Replace("Ruby", "Jade", options);
        Console.WriteLine($"Options replace made {count2} replacements.");

        // Example 3: Regular‑expression replace with a callback to log each replacement.
        FindReplaceOptions regexOptions = new FindReplaceOptions(new ReplacementLogger());
        Regex regex = new Regex(@"\b(\d{4})\b"); // Find any four‑digit number.
        // Replace each year with "2026" while logging the original value.
        int count3 = doc.Range.Replace(regex, "2026", regexOptions);
        Console.WriteLine($"Regex replace made {count3} replacements.");

        // Save the modified document (lifecycle rule: save)
        doc.Save("Output.docx");
    }

    // Callback implementation that logs each replacement.
    private class ReplacementLogger : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            Console.WriteLine($"Replacing \"{args.Match.Value}\" with \"{args.Replacement}\" " +
                              $"at offset {args.MatchOffset} in node type {args.MatchNode.NodeType}.");
            // Perform the default replacement.
            return ReplaceAction.Replace;
        }
    }
}
