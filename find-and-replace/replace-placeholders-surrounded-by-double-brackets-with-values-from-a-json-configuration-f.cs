using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class PlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public PlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The whole match includes the surrounding brackets, e.g. [[Name]]
        string placeholder = args.Match.Value;
        if (placeholder.Length < 4) // Minimum length for [[x]]
            return ReplaceAction.Skip;

        // Extract the key without the brackets.
        string key = placeholder.Substring(2, placeholder.Length - 4);
        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
            return ReplaceAction.Replace;
        }

        // No matching key – leave the placeholder unchanged.
        return ReplaceAction.Skip;
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "FindReplaceExample");
        Directory.CreateDirectory(workDir);

        // Create a JSON configuration file with placeholder values.
        string jsonPath = Path.Combine(workDir, "config.json");
        var configData = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Date", "2023-10-01" },
            { "Location", "New York" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(configData));

        // Create a sample Word document containing placeholders.
        string inputPath = Path.Combine(workDir, "input.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear [[Name]],");
        builder.Writeln("Your appointment is scheduled on [[Date]] at [[Location]].");
        builder.Writeln("Thank you, [[Name]].");
        doc.Save(inputPath);

        // Load the JSON configuration into a dictionary.
        string jsonContent = File.ReadAllText(jsonPath);
        var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent)
                     ?? new Dictionary<string, string>();

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up the find-and-replace options with a callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(values);

        // Regex pattern to match placeholders like [[Key]].
        Regex placeholderPattern = new Regex(@"\[\[(.+?)\]\]");

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace(placeholderPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced. Expected at least one replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(workDir, "output.docx");
        loadedDoc.Save(outputPath);

        // Optional: write a simple report of the operation.
        string reportPath = Path.Combine(workDir, "report.json");
        var report = new
        {
            InputDocument = inputPath,
            OutputDocument = outputPath,
            ReplacementsMade = replacedCount,
            Timestamp = DateTime.UtcNow
        };
        File.WriteAllText(reportPath, JsonConvert.SerializeObject(report, Formatting.Indented));
    }
}
