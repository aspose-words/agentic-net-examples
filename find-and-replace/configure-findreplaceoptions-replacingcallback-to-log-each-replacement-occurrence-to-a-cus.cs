using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that will be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is quick.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up a custom logger that records each replacement occurrence.
        var logger = new ReplaceLogger();

        // Configure FindReplaceOptions to use the logger callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement operation.
        int replacedCount = loaded.Range.Replace("quick", "swift", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output the log to the console.
        Console.WriteLine("Replacements performed:");
        foreach (string entry in logger.Matches)
        {
            Console.WriteLine(entry);
        }

        // Additionally, write the log to a JSON file.
        const string jsonLogPath = "log.json";
        File.WriteAllText(jsonLogPath, JsonConvert.SerializeObject(logger.Matches, Formatting.Indented));
    }

    // Custom logger implementing IReplacingCallback.
    private class ReplaceLogger : IReplacingCallback
    {
        public List<string> Matches { get; } = new List<string>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Record the original matched text, the replacement text, and the offset within the node.
            Matches.Add($"\"{args.Match.Value}\" -> \"{args.Replacement}\" at offset {args.MatchOffset}");
            // Proceed with the default replacement.
            return ReplaceAction.Replace;
        }
    }
}
