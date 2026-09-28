using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PrefixReplacingCallback : IReplacingCallback
{
    // This callback adds the prefix "PRE_" to each matched word.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Build the replacement text.
        string prefixed = "PRE_" + args.Match.Value;
        // Assign the replacement text.
        args.Replacement = prefixed;
        // Indicate that the replacement should be performed.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It contains several words to be prefixed.");
        builder.Writeln("Aspose.Words makes text processing easy.");

        // Save the original document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find-and-replace.
        Document loaded = new Document(inputPath);

        // Define a regex that matches each word (sequence of letters).
        Regex wordRegex = new Regex(@"\b\w+\b", RegexOptions.Compiled);

        // Set up find-and-replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PrefixReplacingCallback();

        // Perform the replacement using the regex and callback.
        // The second argument is a dummy replacement string because the actual
        // replacement text is supplied by the callback.
        int replacedCount = loaded.Range.Replace(wordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
