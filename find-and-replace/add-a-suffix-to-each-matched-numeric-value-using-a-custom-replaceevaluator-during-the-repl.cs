using System;
using System.IO;
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
        builder.Writeln("The package weighs 12 and the box 34.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Set up find-and-replace options with a custom callback that adds a suffix.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericSuffixAppender("kg");

        // Replace every numeric match with the original number plus the suffix.
        int replacedCount = loaded.Range.Replace(new Regex(@"\d+"), string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that appends a suffix to each matched numeric value.
    private class NumericSuffixAppender : IReplacingCallback
    {
        private readonly string _suffix;

        public NumericSuffixAppender(string suffix)
        {
            _suffix = suffix ?? string.Empty;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the suffix to the original numeric match.
            args.Replacement = args.Match.Value + _suffix;
            return ReplaceAction.Replace;
        }
    }
}
