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

        // Configure find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ReplaceWithNodeHandler();

        // Find the placeholder text and replace it with a node.
        doc.Range.Replace(new Regex(@"\[PLACEHOLDER\]"), "", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that replaces each match with a new paragraph node.
    private class ReplaceWithNodeHandler : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Create a new paragraph containing the replacement text.
            Paragraph newParagraph = new Paragraph(args.MatchNode.Document);
            Run run = new Run(args.MatchNode.Document, "Inserted paragraph");
            newParagraph.AppendChild(run);

            // Locate the paragraph that contains the matched text.
            Paragraph matchParagraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);

            // Insert the new paragraph after the matched paragraph.
            CompositeNode parentStory = matchParagraph.ParentNode as CompositeNode;
            parentStory.InsertAfter(newParagraph, matchParagraph);

            // Remove the original paragraph that held the placeholder (optional).
            matchParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }
    }
}
