using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ReplaceWithNodeExample
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Define the pattern to search for (e.g., a placeholder like {INSERT_HERE}).
        Regex placeholderPattern = new Regex(@"\{INSERT_HERE\}");

        // Create FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new InsertNodeCallback());

        // Perform the find-and-replace operation.
        // The replacement string is empty because the callback will handle the insertion.
        doc.Range.Replace(placeholderPattern, string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that inserts a new node (e.g., a paragraph) in place of the matched text.
    private class InsertNodeCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Create the node to insert. Here we insert a new paragraph with some text.
            Paragraph newParagraph = new Paragraph(args.MatchNode.Document);
            Run run = new Run(args.MatchNode.Document, "This is the inserted paragraph.");
            newParagraph.AppendChild(run);

            // The match is inside a Run node; its parent is a Paragraph.
            Paragraph matchParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the new paragraph after the paragraph that contains the match.
            CompositeNode parentStory = matchParagraph.ParentNode;
            parentStory.InsertAfter(newParagraph, matchParagraph);

            // Remove the original paragraph that held the placeholder (optional).
            matchParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }
    }
}
