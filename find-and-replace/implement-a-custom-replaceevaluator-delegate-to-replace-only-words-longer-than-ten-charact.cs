using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class LongWordReplacer : IReplacingCallback
{
    // Called for each regex match found in the document.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched word.
        string word = args.Match.Value;

        // If the word is longer than ten characters, replace it with "[REDACTED]".
        if (word.Length > 10)
        {
            args.Replacement = "[REDACTED]";
            return ReplaceAction.Replace;
        }

        // Otherwise, leave the original word unchanged.
        return ReplaceAction.Skip;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with words of varying lengths.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains someverylongwordindeed and other short words.");
        builder.Writeln("AnotherSuperLongWordHere is also present.");

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Regex to match individual words.
        Regex wordRegex = new Regex(@"\b\w+\b");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new LongWordReplacer();

        // Perform the find-and-replace operation.
        // The replacement string is ignored because the callback decides the actual text.
        int replacedCount = doc.Range.Replace(wordRegex, string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
