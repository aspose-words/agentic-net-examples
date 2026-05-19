using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with numeric values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Invoice numbers: 1001, 1002, 1003.");
        doc.Save("source.docx");

        // Load the document for processing.
        Document loaded = new Document("source.docx");

        // Define the suffix to append to each numeric match.
        const string suffix = "_ID";

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericSuffixCallback(suffix);

        // Replace every numeric value using a regular expression.
        int replacedCount = loaded.Range.Replace(new Regex(@"\d+"), string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that appends a suffix to each matched numeric value.
    private class NumericSuffixCallback : IReplacingCallback
    {
        private readonly string _suffix;

        public NumericSuffixCallback(string suffix) => _suffix = suffix;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the suffix to the original numeric text.
            args.Replacement = args.Match.Value + _suffix;
            return ReplaceAction.Replace;
        }
    }
}
