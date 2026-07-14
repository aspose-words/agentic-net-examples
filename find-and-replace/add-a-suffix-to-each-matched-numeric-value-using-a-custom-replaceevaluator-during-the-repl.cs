using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types

namespace FindAndReplaceSuffixExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with numeric values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample numbers: 123, 456, 789.");
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document for processing.
            Document loadedDoc = new Document(inputPath);

            // Set up find-and-replace options with a custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new NumericSuffixCallback();

            // Use a regular expression to match numeric values.
            Regex numberRegex = new Regex(@"\d+");

            // Perform the replace operation. The replacement string is ignored because the callback provides the actual text.
            int replacementCount = loadedDoc.Range.Replace(numberRegex, string.Empty, options);

            if (replacementCount == 0)
                throw new InvalidOperationException("Expected at least one numeric replacement.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }

        // Callback that appends a suffix to each matched numeric value.
        private class NumericSuffixCallback : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                // Append the desired suffix to the original numeric match.
                args.Replacement = args.Match.Value + "_suffix";
                return ReplaceAction.Replace;
            }
        }
    }
}
