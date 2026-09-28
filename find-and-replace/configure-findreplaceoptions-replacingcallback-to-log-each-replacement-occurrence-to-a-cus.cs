using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplaceLogger : IReplacingCallback
{
    public List<string> Matches { get; } = new List<string>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Log the matched text.
        Matches.Add(args.Match.Value);
        // Perform the replacement.
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
        builder.Writeln("alpha beta alpha gamma alpha");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document.
        Document loaded = new Document(inputPath);

        // Set up the logger callback.
        ReplaceLogger logger = new ReplaceLogger();
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = logger;

        // Perform replacement.
        int replacedCount = loaded.Range.Replace("alpha", "omega", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement.");
        }

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Write the logged matches to a JSON file.
        string logPath = "replacements.json";
        string json = JsonConvert.SerializeObject(logger.Matches, Formatting.Indented);
        File.WriteAllText(logPath, json);

        // Verify that the log file was created.
        if (!File.Exists(logPath))
        {
            throw new InvalidOperationException("Log file was not created.");
        }
    }
}
