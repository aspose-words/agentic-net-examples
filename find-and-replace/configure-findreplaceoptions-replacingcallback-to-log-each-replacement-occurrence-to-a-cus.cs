using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplaceLogger : IReplacingCallback
{
    // Stores each matched text that will be replaced.
    public List<string> Matches { get; } = new List<string>();

    // Called for every match found by the FindReplace engine.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Guard against null Match (should not happen in normal usage).
        if (args?.Match != null)
        {
            Matches.Add(args.Match.Value);
        }

        // Perform the replacement.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with repeated occurrences of the word "alpha".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma alpha");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document (could also continue using the same instance).
        Document loaded = new Document(inputPath);

        // Set up the replacement logger and options.
        ReplaceLogger logger = new ReplaceLogger();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement: replace "alpha" with "omega".
        int replacedCount = loaded.Range.Replace("alpha", "omega", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Serialize the logged matches to JSON and write to a file.
        string jsonLog = JsonConvert.SerializeObject(logger.Matches, Formatting.Indented);
        const string logPath = "replacements.json";
        File.WriteAllText(logPath, jsonLog);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Logged matches saved to: {logPath}");
    }
}
