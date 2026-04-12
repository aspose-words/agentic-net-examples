using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing numeric values.
        builder.Writeln("The shipment includes 15 items, 42 boxes, and 7 pallets.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new NumericSuffixReplacer()
        };

        // Replace every sequence of digits. The actual replacement text is supplied by the callback.
        int replacedCount = doc.Range.Replace(new Regex(@"\d+"), string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No numeric values were found to replace.");

        // Save the modified document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docx");
        doc.Save(outputPath);

        // Inform the user (console output is allowed for logging purposes).
        Console.WriteLine($"Replaced {replacedCount} numeric value(s). Output saved to: {outputPath}");
    }

    // Callback that appends a suffix to each matched numeric value.
    private class NumericSuffixReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the desired suffix (e.g., "kg") to the original numeric match.
            string suffix = "kg";
            args.Replacement = args.Match.Value + suffix;

            // Perform the replacement.
            return ReplaceAction.Replace;
        }
    }
}
