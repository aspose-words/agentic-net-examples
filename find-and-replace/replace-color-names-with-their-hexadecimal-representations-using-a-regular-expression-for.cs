using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing; // Required for HighlightColor
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing color names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue, the grass is green, and the sun is yellow.");
        builder.Writeln("Roses are red, violets are purple, and clouds are white.");
        builder.Writeln("A black cat crossed the path, while a gray mouse ran away.");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Regular expression that matches the desired color names (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|yellow|orange|purple|black|white|gray)\b",
                                      RegexOptions.IgnoreCase);

        // Configure find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions();
        // Highlight the replacement text.
        options.ApplyFont.HighlightColor = Color.Yellow;
        // Use a custom callback to convert the matched color name to its hex value.
        options.ReplacingCallback = new ColorNameHexConverter();

        // Perform the replacement. The replacement string is ignored because the callback supplies it.
        int replacedCount = loaded.Range.Replace(colorRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that converts a matched color name to its hexadecimal representation.
    private class ColorNameHexConverter : IReplacingCallback
    {
        // Mapping from color name (case‑insensitive) to hexadecimal string.
        private static readonly Dictionary<string, string> ColorMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "red",    "#FF0000" },
            { "green",  "#008000" },
            { "blue",   "#0000FF" },
            { "yellow", "#FFFF00" },
            { "orange", "#FFA500" },
            { "purple", "#800080" },
            { "black",  "#000000" },
            { "white",  "#FFFFFF" },
            { "gray",   "#808080" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string colorName = args.Match.Value;
            if (ColorMap.TryGetValue(colorName, out string hex))
                args.Replacement = hex; // Set the replacement text.
            else
                args.Replacement = colorName; // Fallback – should not occur.

            return ReplaceAction.Replace;
        }
    }
}
