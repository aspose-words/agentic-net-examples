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
        builder.Writeln("Invoice numbers: 1001, 1002, 1003.");
        builder.Writeln("Reference IDs: A-200, B-300.");

        // Save the source document locally.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document back from the file system.
        Document loadedDoc = new Document(inputPath);

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericSuffixReplacer("_SUFFIX");

        // Use a regular expression to locate numeric values.
        Regex numericPattern = new Regex(@"\d+");

        // Perform the replace operation. The replacement string is ignored because the callback sets it.
        int replacementCount = loadedDoc.Range.Replace(numericPattern, string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loadedDoc.Save(outputPath);
    }

    // Custom callback that appends a suffix to each matched numeric value.
    private class NumericSuffixReplacer : IReplacingCallback
    {
        private readonly string _suffix;

        public NumericSuffixReplacer(string suffix)
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
