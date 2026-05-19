using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("apple banana apple orange");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Set up find-and-replace options with a custom callback that adds a prefix.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new PrefixReplacer()
        };

        // Perform the replacement for the word "apple".
        int replacedCount = loadedDoc.Range.Replace("apple", string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Output the result count.
        Console.WriteLine($"Replacements made: {replacedCount}");
    }

    // Callback that adds the prefix "PRE_" to each matched word.
    private class PrefixReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            args.Replacement = "PRE_" + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
