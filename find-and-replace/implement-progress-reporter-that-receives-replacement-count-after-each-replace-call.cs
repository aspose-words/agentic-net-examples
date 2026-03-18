using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplacementProgressReporter : IReplacingCallback
{
    private int _replacementCount;

    // Exposes the number of replacements that have occurred so far.
    public int ReplacementCount => _replacementCount;

    // This method is called for every match found by the Replace operation.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        _replacementCount++;

        // Report progress – you can replace this with any other reporting mechanism.
        Console.WriteLine($"Replacement #{_replacementCount}: \"{args.Match.Value}\" → \"{args.Replacement}\" (offset {args.MatchOffset})");

        // Perform the actual replacement.
        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple banana apple. Banana apple banana.");

        // Initialise the progress reporter.
        var reporter = new ReplacementProgressReporter();

        // Configure find/replace options to use the reporter.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = reporter;

        // Perform a case‑insensitive replace of the word "apple" with "orange".
        int totalReplacements = doc.Range.Replace(new Regex("apple", RegexOptions.IgnoreCase), "orange", options);

        // Output the total number of replacements returned by the Replace method.
        Console.WriteLine($"Replace() returned: {totalReplacements}");

        // Output the total number counted by the reporter (should match the above).
        Console.WriteLine($"Reporter counted: {reporter.ReplacementCount}");

        // Save the modified document (optional).
        doc.Save("ProgressReport.docx");
    }
}
