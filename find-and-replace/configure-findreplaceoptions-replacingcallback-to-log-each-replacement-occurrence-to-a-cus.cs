using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma beta alpha");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Set up a logger that records each replacement occurrence.
        var logger = new ReplaceLogger();

        // Configure FindReplaceOptions to use the logger.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("alpha", "omega", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Write the log of replacements to a text file.
        File.WriteAllText("replace_log.txt", logger.GetLog());
    }

    // Custom logger that implements IReplacingCallback.
    private class ReplaceLogger : IReplacingCallback
    {
        private readonly List<string> _matches = new List<string>();
        private readonly StringWriter _logWriter = new StringWriter();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Record the matched text and its location.
            _matches.Add(args.Match.Value);
            _logWriter.WriteLine($"Matched \"{args.Match.Value}\" at offset {args.MatchOffset} in a {args.MatchNode.NodeType} node.");
            // Proceed with the replacement.
            return ReplaceAction.Replace;
        }

        // Returns the accumulated log as a string.
        public string GetLog()
        {
            _logWriter.WriteLine();
            _logWriter.WriteLine("All matches:");
            foreach (var m in _matches)
                _logWriter.WriteLine(m);
            return _logWriter.ToString();
        }
    }
}
