using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text containing color names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The sky is blue, the grass is green, and the sun is yellow.");
        builder.Writeln("Roses are red, violets are purple, and clouds are white.");
        builder.Writeln("A black cat crossed the road while a gray mouse ran away.");

        // Define a regex that matches common color names (case‑insensitive).
        const string pattern = @"\b(red|green|blue|black|white|yellow|orange|purple|gray|grey)\b";

        // Set up find‑replace options with a custom callback that converts each match to its hex value.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ColorNameToHexConverter()
        };

        // Perform the replace operation. The replacement string is ignored because the callback supplies it.
        int replacementCount = doc.Range.Replace(new Regex(pattern, RegexOptions.IgnoreCase), string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No color names were found to replace.");

        // Save the modified document.
        const string outputPath = "ColorNamesReplaced.docx";
        doc.Save(outputPath);

        // Optional: write a simple log to the console.
        Console.WriteLine($"Replaced {replacementCount} color name(s). Output saved to '{outputPath}'.");
    }

    // Callback that replaces a matched color name with its hexadecimal representation.
    private class ColorNameToHexConverter : IReplacingCallback
    {
        // Mapping from color name (lowercase) to hexadecimal string.
        private static readonly Dictionary<string, string> ColorMap = new()
        {
            { "red", "#FF0000" },
            { "green", "#008000" },
            { "blue", "#0000FF" },
            { "black", "#000000" },
            { "white", "#FFFFFF" },
            { "yellow", "#FFFF00" },
            { "orange", "#FFA500" },
            { "purple", "#800080" },
            { "gray", "#808080" },
            { "grey", "#808080" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string colorName = args.Match.Value.ToLowerInvariant();

            if (ColorMap.TryGetValue(colorName, out string hex))
                args.Replacement = hex;
            else
                args.Replacement = args.Match.Value; // Fallback: keep original text.

            return ReplaceAction.Replace;
        }
    }
}
