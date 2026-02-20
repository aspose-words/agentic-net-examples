using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Drawing;

namespace FindReplaceHighlightExample
{
    // Custom callback that highlights each match found during the replace operation.
    public class HighlightReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Apply yellow highlight to the node that contains the match.
            // The match may be inside a Run; ensure we get the Run node.
            var run = args.MatchNode as Run;
            if (run != null)
                run.Font.HighlightColor = Color.Yellow;

            // Keep the original text unchanged.
            args.Replacement = args.Match.Value;

            // Skip the actual replacement because we only wanted to format the existing text.
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX document.
            Document doc = new Document("Input.docx");

            // Set up find/replace options with the custom highlighting callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new HighlightReplacer();

            // Define the text to find. This example uses a regular expression to find the word "sample".
            string pattern = @"\bsample\b";

            // Perform the find/replace operation. The replacement string is the same as the match,
            // because the callback handles the highlighting.
            doc.Range.Replace(new Regex(pattern, RegexOptions.IgnoreCase), "$0", options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
