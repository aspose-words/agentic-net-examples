using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure FindReplaceOptions with a custom callback that inserts a node.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertNodeCallback()
        };

        // Replace the placeholder text "[PLACEHOLDER]" with a node.
        doc.Range.Replace(new Regex(@"\[PLACEHOLDER\]"), string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a new paragraph node instead of performing a text replacement.
    private class InsertNodeCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The document returned by MatchNode.Document is of type DocumentBase.
            // Cast it to Document to be able to create new nodes.
            Document ownerDoc = (Document)args.MatchNode.Document;

            // Create a new paragraph with the desired content.
            Paragraph newParagraph = new Paragraph(ownerDoc);
            Run run = new Run(ownerDoc, "This is the inserted paragraph.");
            newParagraph.AppendChild(run);

            // Insert the new paragraph after the paragraph that contains the match.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            CompositeNode parent = placeholderParagraph.ParentNode;
            parent.InsertAfter(newParagraph, placeholderParagraph);

            // Remove the placeholder paragraph that held the match.
            placeholderParagraph.Remove();

            // Skip the default text replacement because we have handled it.
            return ReplaceAction.Skip;
        }
    }
}
