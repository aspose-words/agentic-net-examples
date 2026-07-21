using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the package list
using Newtonsoft.Json; // Required by the package list

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing several words.
        builder.Writeln("This is a sample document. It contains several words, such as apple, banana, and cherry.");

        // Define a callback that adds a prefix to each matched word.
        IReplacingCallback prefixCallback = new PrefixReplacer("PRE_");

        // Configure find‑replace options to use the callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = prefixCallback
        };

        // Use a regular expression to match individual words.
        Regex wordRegex = new Regex(@"\b\w+\b");

        // Perform the replacement. The callback will set the actual replacement text.
        int replacedCount = doc.Range.Replace(wordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Callback implementation that adds a prefix to each match.
    private class PrefixReplacer : IReplacingCallback
    {
        private readonly string _prefix;

        public PrefixReplacer(string prefix)
        {
            _prefix = prefix ?? string.Empty;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Prepend the prefix to the original matched text.
            args.Replacement = _prefix + args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
