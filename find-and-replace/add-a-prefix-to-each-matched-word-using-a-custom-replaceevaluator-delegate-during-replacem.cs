using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing words to be prefixed.
        builder.Writeln("Apple Banana Cherry");

        // Define the prefix to add to each matched word.
        const string prefix = "PRE_";

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PrefixReplacer(prefix)
        };

        // Use a regular expression to match whole words.
        Regex wordPattern = new Regex(@"\b\w+\b");

        // Perform the replacement. The callback will set the replacement text.
        int replacementCount = doc.Range.Replace(wordPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No words were replaced.");

        // Save the modified document to a local file.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Optional: write the result count to the console.
        Console.WriteLine($"Replacements made: {replacementCount}");
    }

    // Custom callback that adds a prefix to each matched word.
    private class PrefixReplacer : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixReplacer(string prefix)
        {
            _prefix = prefix ?? throw new ArgumentNullException(nameof(prefix));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Prepend the prefix to the original matched text.
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
