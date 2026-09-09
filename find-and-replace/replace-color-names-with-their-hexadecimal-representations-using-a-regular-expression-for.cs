using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with color names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue and the grass is green.");
        builder.Writeln("Red apples, black night, and white snow.");
        builder.Writeln("Gray clouds drift by.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ColorNameHexReplacer()
        };

        // Regular expression to match color names (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|black|white|gray)\b", RegexOptions.IgnoreCase);

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = loaded.Range.Replace(colorRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that converts a matched color name to its hexadecimal representation.
    private class ColorNameHexReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> ColorMap = new()
        {
            { "red",   "#FF0000" },
            { "green", "#008000" },
            { "blue",  "#0000FF" },
            { "black", "#000000" },
            { "white", "#FFFFFF" },
            { "gray",  "#808080" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string matchedName = args.Match.Value.ToLowerInvariant();

            if (ColorMap.TryGetValue(matchedName, out string hex))
            {
                args.Replacement = hex;
                return ReplaceAction.Replace;
            }

            // If the color is not in the map, leave it unchanged.
            return ReplaceAction.Skip;
        }
    }
}
