using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that adds a prefix to each matched word.
    private class PrefixCallback : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixCallback(string prefix) => _prefix = prefix;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Build the replacement text: prefix + original matched word.
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a new blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple banana apple Banana APPLE.");

        const string prefix = "PRE_";

        // Configure find‑replace options: case‑insensitive, whole‑word matches.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true,
            ReplacingCallback = new PrefixCallback(prefix)
        };

        // Perform the replacement for the word "apple".
        // The replacement string is ignored because the callback supplies the actual text.
        int replacedCount = doc.Range.Replace("apple", string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No occurrences of the target word were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the result to the console for verification.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
