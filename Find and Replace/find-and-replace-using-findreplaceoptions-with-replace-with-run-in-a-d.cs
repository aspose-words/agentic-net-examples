using System;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceWithRunHandler : IReplacingCallback
{
    // This method is called for each match found by the Find/Replace engine.
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Create a new Run node that will replace the matched text.
        Run newRun = new Run(args.MatchNode.Document, "Replacement Text");

        // Example: apply custom formatting to the new run.
        newRun.Font.Bold = true;
        newRun.Font.Color = Color.Blue;

        // Insert the new run immediately after the node that contained the match.
        // InsertAfter is defined on CompositeNode, so cast the parent accordingly.
        CompositeNode parent = (CompositeNode)args.MatchNode.ParentNode;
        parent.InsertAfter(newRun, args.MatchNode);

        // Remove the original node that held the matched text.
        args.MatchNode.Remove();

        // Skip the default replacement because we have already performed the custom one.
        return ReplaceAction.Skip;
    }
}

class FindReplaceWithRunExample
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceWithRunHandler()
        };

        // Perform the find-and-replace operation.
        // The pattern "PLACEHOLDER" will be replaced by a new Run node.
        doc.Range.Replace(new Regex("PLACEHOLDER"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
