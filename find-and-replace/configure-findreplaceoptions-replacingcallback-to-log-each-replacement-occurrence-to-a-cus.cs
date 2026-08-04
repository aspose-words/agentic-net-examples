using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ReplaceLogger : IReplacingCallback
{
    private readonly StringBuilder _log = new StringBuilder();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        _log.AppendLine($"Match \"{args.Match.Value}\" at offset {args.MatchOffset} in node {args.MatchNode.NodeType}");
        return ReplaceAction.Replace;
    }

    public string GetLog() => _log.ToString();
}

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string logPath = "replace_log.txt";

        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure the find‑replace options with a custom logger.
        ReplaceLogger logger = new ReplaceLogger();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("alpha", "omega", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Write the replacement log to a file.
        File.WriteAllText(logPath, logger.GetLog());
    }
}
