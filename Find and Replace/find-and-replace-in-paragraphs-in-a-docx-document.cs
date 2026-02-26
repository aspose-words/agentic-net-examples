using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        // The Document constructor handles opening the file – no custom creation code needed.
        Document doc = new Document(@"Input.docx");

        // Example 1: Simple literal replace.
        // Replaces every occurrence of the placeholder "_FullName_" with "John Doe".
        int count1 = doc.Range.Replace("_FullName_", "John Doe");
        Console.WriteLine($"Literal replacements made: {count1}");

        // Example 2: Case‑sensitive replace with options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true,               // Respect case while searching.
            FindWholeWordsOnly = true       // Replace only whole word matches.
        };
        int count2 = doc.Range.Replace("Ruby", "Jade", options);
        Console.WriteLine($"Case‑sensitive replacements made: {count2}");

        // Example 3: Regular‑expression replace that also changes paragraph alignment.
        // Replace a period followed by a paragraph break with an exclamation point,
        // and right‑align every paragraph that contained the match.
        FindReplaceOptions regexOptions = new FindReplaceOptions();
        regexOptions.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;
        int count3 = doc.Range.Replace(@".&p", "!&p", regexOptions);
        Console.WriteLine($"Regex replacements made (with alignment change): {count3}");

        // Example 4: Using a callback to log each replacement.
        FindReplaceOptions callbackOptions = new FindReplaceOptions(new ReplacementLogger());
        doc.Range.Replace(new Regex(@"New York City|NYC"), "Washington", callbackOptions);

        // Save the modified document.
        // The Save method follows the required lifecycle rule.
        doc.Save(@"Output.docx");
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
