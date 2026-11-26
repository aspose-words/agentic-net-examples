using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with numeric values in US format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The total is 1,234.56 dollars.");
        builder.Writeln("Another value: 12,345.67 and 123.45.");
        builder.Writeln("No change needed for 1234,56.");

        // Define a regular expression that matches numbers like 1,234.56 or 12,345.
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find‑replace options with a custom callback to localize the format.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericFormatReplacer();

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(numberRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric format replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Replacements performed: {replacedCount}. Output saved to '{outputPath}'.");
    }
}

// Callback that converts US numeric format (1,234.56) to European format (1.234,56).
public class NumericFormatReplacer : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        string original = args.Match.Value;

        // Preserve numbers without a decimal separator (e.g., "1234,56") unchanged.
        if (!original.Contains(",") && !original.Contains("."))
        {
            args.Replacement = original;
            return ReplaceAction.Replace;
        }

        // Swap commas and periods.
        string temp = original.Replace(",", "__COMMA__");
        string swapped = temp.Replace(".", ",");
        string result = swapped.Replace("__COMMA__", ".");

        args.Replacement = result;
        return ReplaceAction.Replace;
    }
}
