using System;
using System.Collections.Generic;
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
        builder.Writeln("The sky is blue, the grass is green, and the sun is yellow.");
        builder.Writeln("My favorite colors are red, light gray and dark red.");
        builder.Writeln("Black and white are also classic choices.");

        // Save the source document (optional, just to illustrate the workflow).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Prepare a regex that matches the color names (case‑insensitive).
        const string pattern = @"\b(?:light\s+gray|dark\s+red|red|green|blue|yellow|black|white|gray)\b";
        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ColorHexConverter()
        };

        // Perform the replacement. The replacement string argument is ignored when a callback is used.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that converts a matched color name to its hexadecimal representation.
    private class ColorHexConverter : IReplacingCallback
    {
        // Mapping from normalized color name to hex string.
        private static readonly Dictionary<string, string> ColorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "red", "#FF0000" },
            { "green", "#008000" },
            { "blue", "#0000FF" },
            { "yellow", "#FFFF00" },
            { "black", "#000000" },
            { "white", "#FFFFFF" },
            { "gray", "#808080" },
            { "light gray", "#D3D3D3" },
            { "dark red", "#8B0000" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Normalise the matched value (trim and collapse multiple spaces).
            string key = Regex.Replace(args.Match.Value.Trim(), @"\s+", " ").ToLowerInvariant();

            if (ColorMap.TryGetValue(key, out string hex))
            {
                args.Replacement = hex;
                return ReplaceAction.Replace;
            }

            // If the color is not in the map, leave it unchanged.
            return ReplaceAction.Skip;
        }
    }
}
