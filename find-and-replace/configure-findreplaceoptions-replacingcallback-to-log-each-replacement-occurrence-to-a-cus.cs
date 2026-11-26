using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementLogger : IReplacingCallback
{
    public List<string> Matches { get; } = new List<string>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Log the original text that is being replaced.
        Matches.Add(args.Match.Value);
        // Perform the replacement as usual.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog. The fox is clever.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up the custom logger callback.
        ReplacementLogger logger = new ReplacementLogger();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the find-and-replace operation.
        int replacedCount = loaded.Range.Replace("fox", "cat", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Serialize the logged matches to a JSON file.
        string jsonLog = JsonConvert.SerializeObject(logger.Matches, Formatting.Indented);
        const string logPath = "replacements.json";
        File.WriteAllText(logPath, jsonLog);

        // Verify that the log file was created.
        if (!File.Exists(logPath))
            throw new InvalidOperationException("Replacement log file was not created.");

        // Optionally, write a brief confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Log written to: {Path.GetFullPath(logPath)}");
    }
}
