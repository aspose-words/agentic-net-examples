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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue.");
        builder.Writeln("The grass is green.");
        builder.Writeln("The rose is red.");
        builder.Writeln("Sunlight is yellow.");
        builder.Writeln("Night is black.");
        builder.Writeln("Snow is white.");
        builder.Writeln("A violet flower is purple.");
        builder.Writeln("A bright orange fruit.");

        // Save the input document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define a regex that matches the color names we want to replace (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|yellow|orange|purple|black|white)\b",
                                     RegexOptions.IgnoreCase);

        // Set up the replacement callback that converts a color name to its hex code.
        var callback = new ColorNameToHexCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement using the regex and the callback.
        int replacedCount = doc.Range.Replace(colorRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Serialize the replacement report to JSON.
        string jsonReport = JsonConvert.SerializeObject(callback.Replacements, Formatting.Indented);
        const string reportPath = "report.json";
        File.WriteAllText(reportPath, jsonReport);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("The replacement report was not created.");
    }
}

// Holds information about a single replacement.
public class ReplacementInfo
{
    public string Original { get; set; } = string.Empty;
    public string Hex { get; set; } = string.Empty;
}

// Callback that maps color names to hexadecimal values.
public class ColorNameToHexCallback : IReplacingCallback
{
    private static readonly Dictionary<string, string> ColorMap = new(StringComparer.OrdinalIgnoreCase)
    {
        { "red",    "#FF0000" },
        { "green",  "#00FF00" },
        { "blue",   "#0000FF" },
        { "yellow", "#FFFF00" },
        { "orange", "#FFA500" },
        { "purple", "#800080" },
        { "black",  "#000000" },
        { "white",  "#FFFFFF" }
    };

    public List<ReplacementInfo> Replacements { get; } = new();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        string original = args.Match.Value;
        if (ColorMap.TryGetValue(original, out string hex))
        {
            args.Replacement = hex;
            Replacements.Add(new ReplacementInfo { Original = original, Hex = hex });
        }
        else
        {
            // If the color is not in the map, keep the original text.
            args.Replacement = original;
        }

        return ReplaceAction.Replace;
    }
}
