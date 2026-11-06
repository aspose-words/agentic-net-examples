using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that adds a prefix to each matched word.
    private class PrefixReplacer : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixReplacer(string prefix) => _prefix = prefix;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Build the replacement text with the desired prefix.
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. It contains several words to replace.");

        // Save the sample document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches each word.
        Regex wordRegex = new Regex(@"\b\w+\b", RegexOptions.Compiled);

        // Prefix to add to each matched word.
        const string prefix = "PRE_";

        // Set up FindReplaceOptions with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PrefixReplacer(prefix);

        // Perform the replacement using the regex and callback.
        // The overload requires a replacement string; it will be ignored because the callback sets args.Replacement.
        int replacedCount = loaded.Range.Replace(wordRegex, string.Empty, options);

        // Validate that replacements occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No words were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replaced {replacedCount} words. Output saved to '{outputPath}'.");
    }
}
