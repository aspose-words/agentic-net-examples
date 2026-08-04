using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsFindAndReplaceExample
{
    // Callback that replaces only words longer than ten characters.
    public class LongWordReplacer : IReplacingCallback
    {
        // Called for each match found by the replace operation.
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // If the matched word length exceeds ten characters, replace it.
            if (args.Match.Value.Length > 10)
            {
                // Replace with a placeholder text.
                args.Replacement = "LONGWORD";
                return ReplaceAction.Replace;
            }

            // Otherwise, skip this match and leave the original text unchanged.
            return ReplaceAction.Skip;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with words of varying lengths.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Short words and extraordinarilylongword should be processed.");
            builder.Writeln("AnotherSuperLongWord appears here, while tiny stays.");
            // Save the source document.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document for processing.
            Document loadedDoc = new Document(inputPath);

            // Define a regex that matches whole words.
            Regex wordRegex = new Regex(@"\b\w+\b", RegexOptions.Compiled);

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new LongWordReplacer()
            };

            // Perform the replace operation. The replacement string is ignored because the callback sets it.
            int replacementCount = loadedDoc.Range.Replace(wordRegex, string.Empty, options);

            // Ensure that at least one replacement occurred.
            if (replacementCount == 0)
                throw new InvalidOperationException("Expected at least one word longer than ten characters to be replaced.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
