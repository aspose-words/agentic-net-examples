using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders surrounded by percent signs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello %USERNAME%!");
        builder.Writeln("Your home directory is %HOME%.");
        builder.Writeln("Undefined variable: %UNDEFINED_VAR%.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define a regex that matches placeholders like %PLACEHOLDER%.
        Regex placeholderRegex = new Regex("%[A-Za-z0-9_]+%");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EnvironmentVariableReplacer()
        };

        // Perform the replace operation. The replacement string is ignored because the callback supplies it.
        int replacedCount = loadedDoc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Optional: write a JSON report of the replacements performed.
        var report = ((EnvironmentVariableReplacer)options.ReplacingCallback).GetReport();
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("replacement_report.json", jsonReport);
    }
}

// Custom callback that replaces each %PLACEHOLDER% with the corresponding environment variable value.
public class EnvironmentVariableReplacer : IReplacingCallback
{
    private readonly List<ReplacementInfo> _replacements = new List<ReplacementInfo>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the placeholder name without the surrounding percent signs.
        string placeholder = args.Match.Value.Trim('%');
        // Retrieve the environment variable value.
        string envValue = Environment.GetEnvironmentVariable(placeholder) ?? string.Empty;

        // Record the replacement details.
        _replacements.Add(new ReplacementInfo
        {
            Placeholder = args.Match.Value,
            EnvironmentVariable = placeholder,
            ReplacementValue = envValue
        });

        // Set the replacement text.
        args.Replacement = envValue;
        return ReplaceAction.Replace;
    }

    // Returns a report that can be serialized to JSON.
    public IReadOnlyList<ReplacementInfo> GetReport() => _replacements;
}

// Simple DTO for reporting each replacement.
public class ReplacementInfo
{
    public string Placeholder { get; set; } = string.Empty;
    public string EnvironmentVariable { get; set; } = string.Empty;
    public string ReplacementValue { get; set; } = string.Empty;
}
