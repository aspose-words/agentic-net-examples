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
        builder.Writeln("Apple banana apple Banana APPLE.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up a callback that adds a prefix to each matched word.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false, // Case‑insensitive search.
            ReplacingCallback = new PrefixAddingCallback("Fruit_")
        };

        // Replace all occurrences of the word "apple" using the callback.
        int replacedCount = loadedDoc.Range.Replace("apple", string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Inform the user about the operation.
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Modified document saved as: {outputPath}");
    }

    // Callback that prefixes each matched word with a custom string.
    private class PrefixAddingCallback : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixAddingCallback(string prefix) => _prefix = prefix;

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
