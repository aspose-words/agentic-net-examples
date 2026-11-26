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
        // Create a sample document with color names.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue and the grass is green. The rose is red.");
        doc.Save("input.docx");

        // Load the document to perform find-and-replace.
        var loaded = new Document("input.docx");

        // Mapping from color names to hexadecimal values.
        var colorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "red", "#FF0000" },
            { "green", "#00FF00" },
            { "blue", "#0000FF" }
        };

        // Regular expression to match the color names as whole words, case‑insensitive.
        var regex = new Regex(@"\b(red|green|blue)\b", RegexOptions.IgnoreCase);

        // Callback that replaces each matched color name with its hex code.
        var replacer = new ColorNameReplacer(colorMap);
        var options = new FindReplaceOptions { ReplacingCallback = replacer };

        // Perform the replacement. The replacement string argument is ignored when a callback is used.
        int replacedCount = loaded.Range.Replace(regex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Write a JSON report of the performed replacements.
        string jsonReport = JsonConvert.SerializeObject(replacer.Replacements, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}

// Callback implementation for color name replacement.
public class ColorNameReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _colorMap;
    public List<ReplacementInfo> Replacements { get; } = new List<ReplacementInfo>();

    public ColorNameReplacer(IDictionary<string, string> colorMap)
    {
        _colorMap = colorMap ?? throw new ArgumentNullException(nameof(colorMap));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        string original = args.Match.Value;
        if (_colorMap.TryGetValue(original, out string hex))
        {
            args.Replacement = hex;
            Replacements.Add(new ReplacementInfo { Original = original, Replacement = hex });
            return ReplaceAction.Replace;
        }

        // No mapping found; leave the original text unchanged.
        return ReplaceAction.Skip;
    }
}

// Simple DTO for JSON reporting.
public class ReplacementInfo
{
    public string Original { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
}
