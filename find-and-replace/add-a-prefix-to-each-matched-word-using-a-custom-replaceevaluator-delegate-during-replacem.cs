using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PrefixReplacer : IReplacingCallback
{
    private readonly string _prefix;

    public PrefixReplacer(string prefix)
    {
        _prefix = prefix;
    }

    // Called for each match found by Range.Replace.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Prepend the desired prefix to the matched text.
        args.Replacement = _prefix + args.Match.Value;
        // Perform the replacement.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma alpha");

        // Define the prefix to add to each matched word.
        const string prefix = "pre-";

        // Set up find/replace options and assign the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PrefixReplacer(prefix);

        // Perform the replacement for the word "alpha" using the callback.
        // The second argument (replacement text) is ignored when a callback is used,
        // so we pass an empty string.
        int replacedCount = doc.Range.Replace("alpha", string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
