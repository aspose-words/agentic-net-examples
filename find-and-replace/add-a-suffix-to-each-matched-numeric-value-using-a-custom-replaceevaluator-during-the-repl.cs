using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing numeric values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The order numbers are 1001, 1002, and 1003.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression that matches numeric values.
        Regex numericRegex = new Regex(@"\d+");

        // Suffix to append to each numeric match.
        const string suffix = "_suffix";

        // Set up find‑replace options with a custom callback that appends the suffix.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new AppendSuffixCallback(suffix);

        // Perform the replacement. The replacement string argument is ignored because the callback supplies the new text.
        int replacedCount = loaded.Range.Replace(numericRegex, string.Empty, options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Custom callback that appends a suffix to each regex match.
    private class AppendSuffixCallback : IReplacingCallback
    {
        private readonly string _suffix;

        public AppendSuffixCallback(string suffix)
        {
            _suffix = suffix;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the suffix to the matched text.
            args.Replacement = args.Match.Value + _suffix;
            return ReplaceAction.Replace;
        }
    }
}
