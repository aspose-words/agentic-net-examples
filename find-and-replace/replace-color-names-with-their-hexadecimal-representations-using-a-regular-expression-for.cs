using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Used for color name handling

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write sample text containing color names.
        builder.Writeln("The sky is blue, the grass is green, and the fire is red.");
        builder.Writeln("A light gray cloud floats above a black night.");
        builder.Writeln("White snow covers the ground.");

        // Define a regex that matches the color names we want to replace.
        const string pattern = @"\b(red|green|blue|light gray|gray|black|white)\b";
        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ColorNameToHexConverter()
        };

        // Perform the replacement. The callback will supply the replacement text.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No color names were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Write the result count to the console.
        Console.WriteLine($"Replaced {replacedCount} color name(s). Output saved to '{outputPath}'.");
    }

    // Callback that converts a matched color name to its hexadecimal representation.
    private class ColorNameToHexConverter : IReplacingCallback
    {
        // Mapping from color name (case‑insensitive) to hexadecimal string.
        private static readonly Dictionary<string, string> ColorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "red", "#FF0000" },
            { "green", "#008000" },
            { "blue", "#0000FF" },
            { "light gray", "#D3D3D3" },
            { "gray", "#808080" },
            { "black", "#000000" },
            { "white", "#FFFFFF" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Get the matched color name.
            string colorName = args.Match.Value.Trim();

            // Look up the hexadecimal value; if not found, keep the original text.
            if (ColorMap.TryGetValue(colorName, out string hex))
                args.Replacement = hex;
            else
                args.Replacement = colorName;

            // Demonstrate usage of Aspose.Drawing without assigning to System.Drawing types.
            // For example, we could log the Aspose.Drawing color (not required for the task).
            // Aspose.Drawing.Color asposeColor = Color.FromName(colorName);
            // Console.WriteLine($"Aspose.Drawing color for '{colorName}': {asposeColor}");

            return ReplaceAction.Replace;
        }
    }
}
