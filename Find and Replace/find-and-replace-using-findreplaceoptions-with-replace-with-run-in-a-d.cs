using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with a custom callback that inserts a Run node.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new RunReplacingHandler()
        };

        // Perform the replace. The pattern to find is "_FullName_".
        // The replacement string is left empty because the callback will handle insertion.
        doc.Range.Replace("_FullName_", string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Custom callback that replaces the found text with a new Run node.
    private class RunReplacingHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Create a new Run containing the desired replacement text.
            Run newRun = new Run(args.MatchNode.Document, "John Doe")
            {
                // Example of applying formatting to the new Run.
                Font = { Bold = true, Size = 12 }
            };

            // Insert the new Run after the node where the match starts.
            CompositeNode parent = args.MatchNode.ParentNode as CompositeNode;
            parent?.InsertAfter(newRun, args.MatchNode);

            // Remove all nodes that were part of the original match.
            // The match may span multiple nodes, so iterate from the start to the end node.
            Node current = args.MatchNode;
            while (current != null && current != args.MatchEndNode)
            {
                Node next = current.NextSibling;
                current.Remove();
                current = next;
            }

            // Remove the final node of the match.
            args.MatchEndNode?.Remove();

            // Skip the default replacement because we have already performed the custom insertion.
            return ReplaceAction.Skip;
        }
    }
}
