using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new RunReplacingHandler()
        };

        // Replace the placeholder "[PLACEHOLDER]" with a new Run node.
        // The replacement string is ignored because we handle the replacement in the callback.
        doc.Range.Replace(new Regex(@"\[PLACEHOLDER\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that replaces each match with a new Run containing custom text.
    private class RunReplacingHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Create a Run with the desired replacement text and formatting.
            Run newRun = new Run(args.MatchNode.Document, "New Text")
            {
                Font = { Bold = true, Size = 12 }
            };

            // The node that contains the match is usually a Run. Its parent is a CompositeNode
            // (e.g., Paragraph, TableCell, etc.) which provides the InsertAfter method.
            CompositeNode parent = args.MatchNode.ParentNode as CompositeNode;
            if (parent != null)
            {
                parent.InsertAfter(newRun, args.MatchNode);
            }

            // Remove the original placeholder node.
            args.MatchNode.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }
    }
}
