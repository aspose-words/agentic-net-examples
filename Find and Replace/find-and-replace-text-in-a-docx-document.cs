using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindAndReplaceExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        // The Document constructor automatically detects the file format.
        Document doc = new Document("InputDocument.docx");

        // Define the text to find and the replacement text.
        string findText = "OldValue";
        string replaceText = "NewValue";

        // Simple literal replace (case‑insensitive, whole word not required).
        doc.Range.Replace(findText, replaceText, new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        });

        // Alternatively, use a regular expression with custom options.
        // Example: replace any occurrence of "Item[0-9]+" with "Product".
        Regex regex = new Regex(@"Item\d+", RegexOptions.IgnoreCase);
        FindReplaceOptions regexOptions = new FindReplaceOptions
        {
            // Example of a callback that could modify each replacement.
            ReplacingCallback = new CustomReplacingCallback()
        };
        doc.Range.Replace(regex, "Product", regexOptions);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }

    // Optional callback to demonstrate how to intervene during replacement.
    private class CustomReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Log the original match (could be written to console, file, etc.).
            Console.WriteLine($"Replacing \"{args.Match.Value}\" at offset {args.MatchOffset}.");

            // Optionally modify the replacement text.
            // args.Replacement = $"[Replaced:{args.Match.Value}]";

            // Proceed with the replacement.
            return ReplaceAction.Replace;
        }
    }
}
