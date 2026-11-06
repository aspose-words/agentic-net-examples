using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as per package requirement

public class Program
{
    public static void Main()
    {
        // Create a sample document with color names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue and the grass is green.");
        builder.Writeln("Roses are red, violets are blue, and the sun is yellow.");
        builder.Writeln("Black cats and white dogs.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches common color names (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|yellow|black|white)\b", RegexOptions.IgnoreCase);

        // Set up the callback that will replace each color name with its hex code.
        var callback = new ColorNameToHexCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement. The replacement string is ignored because the callback supplies the value.
        int replacedCount = loaded.Range.Replace(colorRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one color name replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple JSON report of the replacements performed.
        string reportPath = "replacementReport.json";
        string jsonReport = JsonConvert.SerializeObject(callback.Replacements, Formatting.Indented);
        File.WriteAllText(reportPath, jsonReport);
    }
}

// Callback that converts matched color names to their hexadecimal representations.
public class ColorNameToHexCallback : IReplacingCallback
{
    private static readonly Dictionary<string, string> ColorMap = new(StringComparer.OrdinalIgnoreCase)
    {
        { "red", "#FF0000" },
        { "green", "#008000" },
        { "blue", "#0000FF" },
        { "yellow", "#FFFF00" },
        { "black", "#000000" },
        { "white", "#FFFFFF" }
    };

    // Stores each original color name and its replacement for reporting.
    public List<ReplacementRecord> Replacements { get; } = new();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Ensure the match is not null.
        if (args.Match is null)
            return ReplaceAction.Skip;

        string original = args.Match.Value;
        if (!ColorMap.TryGetValue(original, out string? hex))
            return ReplaceAction.Skip; // No mapping found; skip replacement.

        // Set the replacement text.
        args.Replacement = hex;

        // Record the replacement.
        Replacements.Add(new ReplacementRecord { Original = original, Replacement = hex });

        return ReplaceAction.Replace;
    }
}

// Simple DTO for JSON reporting.
public class ReplacementRecord
{
    public string Original { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
}
