using System;
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
        builder.Writeln("Order numbers: 123, 456, and 789.");
        builder.Writeln("Reference IDs: 1001, 1002.");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericSuffixReplacer();

        // Use a regular expression to match one or more digits.
        Regex numberRegex = new Regex(@"\d+");

        // Perform the replace operation. The callback will add a suffix to each match.
        int replacedCount = doc.Range.Replace(numberRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        doc.Save("output.docx");

        // Optional: output the resulting text to the console.
        Console.WriteLine("Replaced text:");
        Console.WriteLine(doc.GetText().Trim());
    }

    // Custom callback that appends a suffix to each numeric match.
    private class NumericSuffixReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the suffix "_suffix" to the original numeric value.
            args.Replacement = args.Match.Value + "_suffix";
            return ReplaceAction.Replace;
        }
    }
}
