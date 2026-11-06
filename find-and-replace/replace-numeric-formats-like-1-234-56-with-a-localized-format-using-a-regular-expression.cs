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
        builder.Writeln("Another amount: 12,345.67 euros.");
        builder.Writeln("Simple number: 123.45.");
        builder.Writeln("No thousand separator: 987.65.");
        // Save the original document (optional, just for demonstration).
        doc.Save("input.docx");

        // Define a regex that matches numbers with optional thousand separators and a decimal part.
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Set up find/replace options with a custom callback to perform locale conversion.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericLocaleCallback();

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(numberRegex, "", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that converts a matched US-formatted number to a European-style format.
    private class NumericLocaleCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The original matched text, e.g., "1,234.56".
            string original = args.Match.Value;

            // Swap commas and periods: commas become temporary placeholder, periods become commas,
            // then placeholder becomes periods.
            string temp = original.Replace(",", "#");
            string swapped = temp.Replace(".", ",");
            string localized = swapped.Replace("#", ".");

            // Set the replacement text.
            args.Replacement = localized;

            return ReplaceAction.Replace;
        }
    }
}
