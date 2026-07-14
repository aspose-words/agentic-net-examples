using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world. This is a sample document with several words.");

        // Save the initial document (optional, just to demonstrate file I/O).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Set up a custom callback that adds a prefix to each matched word.
        const string prefix = "PRE_";
        PrefixCallback callback = new PrefixCallback(prefix);
        FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = callback };

        // Define a regular expression that matches individual words.
        Regex wordRegex = new Regex(@"\b\w+\b");

        // Perform the replace operation. The replacement string is ignored because the callback sets it.
        int replacedCount = loaded.Range.Replace(wordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Custom callback that prefixes each matched word.
    private class PrefixCallback : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixCallback(string prefix)
        {
            _prefix = prefix ?? string.Empty;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Add the prefix to the original matched text.
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
