using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains severalwordswithmorethan10characters and short words.");
        builder.Writeln("Anotherlongwordexample should be replaced, while tiny stays.");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Set up a callback that replaces only words longer than ten characters.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LongWordReplacer()
        };

        // Use a regex that matches whole words.
        Regex wordRegex = new Regex(@"\b\w+\b");

        // Perform the replace operation.
        int replacedCount = loaded.Range.Replace(wordRegex, string.Empty, options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that replaces a match only if its length exceeds ten characters.
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // If the matched word is longer than 10 characters, replace it with "SHORT".
            if (args.Match.Value.Length > 10)
            {
                args.Replacement = "SHORT";
                return ReplaceAction.Replace;
            }

            // Otherwise, skip this match.
            return ReplaceAction.Skip;
        }
    }
}
