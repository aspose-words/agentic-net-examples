using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with HTTP hyperlinks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with hyperlinks:");
        builder.InsertHyperlink("Example Site", "http://example.com/page1", true);
        builder.Writeln();
        builder.InsertHyperlink("Test Site", "http://test.com", true);
        builder.Writeln();

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches an HTTP URL.
        Regex httpUrlPattern = new Regex(@"http://\S+", RegexOptions.IgnoreCase);

        // Set up find-and-replace options with a custom callback that changes the scheme to HTTPS.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new SchemeReplacer();

        // Perform the replacement. The replacement string argument is ignored because the callback provides the new value.
        int replacedCount = loaded.Range.Replace(httpUrlPattern, string.Empty, options);

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one hyperlink scheme replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replaced {replacedCount} hyperlink scheme(s). Output saved to '{outputPath}'.");
    }

    // Callback that replaces the "http://" prefix with "https://".
    private class SchemeReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace only the scheme part of the matched URL.
            args.Replacement = args.Match.Value.Replace("http://", "https://", StringComparison.OrdinalIgnoreCase);
            return ReplaceAction.Replace;
        }
    }
}
