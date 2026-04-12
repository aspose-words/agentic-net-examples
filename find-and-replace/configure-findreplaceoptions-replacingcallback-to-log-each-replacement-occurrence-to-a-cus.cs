using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FindReplaceWithLogging
{
    public static void Main()
    {
        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a test. Hello World!");

        // Configure find-and-replace options with a custom logger.
        FindReplaceOptions options = new FindReplaceOptions();
        ReplacementLogger logger = new ReplacementLogger();
        options.ReplacingCallback = logger;

        // Perform the replacement using a regular expression.
        int replacementCount = doc.Range.Replace(new Regex("Hello"), "Hi", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Write the log to console and to a file.
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine("Replacement log:");
        Console.WriteLine(logger.GetLog());

        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "ReplacementLog.txt");
        File.WriteAllText(logPath, logger.GetLog());
    }

    // Custom logger that records each replacement occurrence.
    private class ReplacementLogger : IReplacingCallback
    {
        private readonly StringBuilder _log = new StringBuilder();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            _log.AppendLine(
                $"\"{args.Match.Value}\" replaced with \"{args.Replacement}\" " +
                $"at offset {args.MatchOffset} in node type {args.MatchNode.NodeType}.");
            // Allow the replacement to proceed.
            return ReplaceAction.Replace;
        }

        public string GetLog()
        {
            return _log.ToString();
        }
    }
}
