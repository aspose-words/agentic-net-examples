using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that appends a suffix to each numeric match.
    private class SuffixAppendingCallback : IReplacingCallback
    {
        private readonly string _suffix;

        public SuffixAppendingCallback(string suffix)
        {
            _suffix = suffix;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Append the suffix to the original matched value.
            args.Replacement = args.Match.Value + _suffix;
            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a new document and write sample text containing numeric values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Invoice 1001 total is 2500 dollars.");
        builder.Writeln("Reference numbers: 12345, 67890.");

        // Regex that matches one or more digits.
        Regex numberRegex = new Regex(@"\d+");

        // Set up find‑replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new SuffixAppendingCallback("_SUFFIX");

        // Perform the replace operation using the regex and callback.
        int replacedCount = doc.Range.Replace(numberRegex, string.Empty, options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
