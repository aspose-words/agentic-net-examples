using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ColorNameToHexConverter : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Get the matched color name (e.g., "Red").
        string colorName = args.Match.Value;

        // Convert the name to a System.Drawing.Color.
        Color color = Color.FromName(colorName);

        // Build the hexadecimal representation "#RRGGBB".
        string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

        // Set the replacement text.
        args.Replacement = hex;

        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Load or create a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text containing color names.
        builder.Writeln("The sky is blue, the grass is green, and the sun is yellow.");
        builder.Writeln("My favorite colors are Red, Black, and White.");

        // Set up find/replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ColorNameToHexConverter();

        // Regular expression to match common color names (case‑insensitive).
        Regex colorRegex = new Regex(@"\b(red|green|blue|yellow|orange|purple|black|white|gray|grey)\b",
                                      RegexOptions.IgnoreCase);

        // Perform the replacement.
        doc.Range.Replace(colorRegex, "", options);

        // Save the modified document.
        doc.Save("ColorNamesConverted.docx");
    }
}
