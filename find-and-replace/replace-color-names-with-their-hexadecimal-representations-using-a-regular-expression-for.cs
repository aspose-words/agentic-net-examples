using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some color names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue, the grass is green, and the fire is red.");
        builder.Writeln("A light gray cat sits on a black mat.");
        builder.Writeln("Sunset colors: orange, pink, and yellow.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression to find color names (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|light gray|gray|black|white|yellow|orange|purple|pink)\b",
                                      RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ColorHexReplacer()
        };

        // Perform the replacement. The callback supplies the hexadecimal value.
        int replacedCount = loaded.Range.Replace(colorRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that converts a matched color name to its hexadecimal representation.
    private class ColorHexReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> ColorMap = new()
        {
            { "red", "#FF0000" },
            { "green", "#008000" },
            { "blue", "#0000FF" },
            { "light gray", "#D3D3D3" },
            { "gray", "#808080" },
            { "black", "#000000" },
            { "white", "#FFFFFF" },
            { "yellow", "#FFFF00" },
            { "orange", "#FFA500" },
            { "purple", "#800080" },
            { "pink", "#FFC0CB" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Normalise the matched value for dictionary lookup.
            string key = args.Match.Value.ToLowerInvariant();

            if (ColorMap.TryGetValue(key, out string hex))
                args.Replacement = hex;
            else
                args.Replacement = args.Match.Value; // Fallback – should not happen.

            return ReplaceAction.Replace;
        }
    }
}
