using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our new location in New York City is opening tomorrow. Hope to see all our NYC-based customers at the opening!");

        // Instantiate the custom logger.
        var logger = new ReplacementLogger();

        // Configure FindReplaceOptions to use the logger as the callback.
        FindReplaceOptions options = new FindReplaceOptions(logger);

        // Perform the replace operation; the logger will be invoked for each match.
        doc.Range.Replace(new Regex("New York City|NYC"), "Washington", options);

        // Output the modified document text.
        Console.WriteLine(doc.GetText().Trim());

        // Output the log of replacement occurrences.
        Console.WriteLine(logger.GetLog().Trim());
    }

    // Custom logger that records each replacement occurrence.
    private class ReplacementLogger : IReplacingCallback
    {
        // Called before each replacement.
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Log original match, replacement text, offset and node type.
            _log.AppendLine($"\"{args.Match.Value}\" -> \"{args.Replacement}\" at offset {args.MatchOffset} in {args.MatchNode.NodeType} node.");
            // Continue with the default replacement action.
            return ReplaceAction.Replace;
        }

        // Retrieve the accumulated log.
        public string GetLog()
        {
            return _log.ToString();
        }

        private readonly StringBuilder _log = new StringBuilder();
    }
}
