using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that replaces only words longer than ten characters with "SHORT".
    private class LongWordReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // If the matched word is longer than ten characters, replace it.
            if (args.Match.Value.Length > 10)
                args.Replacement = "SHORT";

            // Apply the (possibly modified) replacement.
            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(
            "This document contains some extraordinarilylongword and anotherSupercalifragilisticexpialidocious example.");
        builder.Writeln("Short words stay unchanged.");

        // Regex that matches words longer than ten characters.
        Regex longWordRegex = new Regex(@"\b\w{11,}\b", RegexOptions.Compiled);

        // Set up find/replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LongWordReplacer()
        };

        // Perform the replace using the regex and callback.
        // The replacement string is ignored because the callback supplies the value.
        int replacedCount = doc.Range.Replace(longWordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replacements performed: {replacedCount}. Output saved to '{outputPath}'.");
    }
}
