using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsReplaceWithNode
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("Input.docx");

            // Define the placeholder to search for.
            string placeholder = @"\[PLACEHOLDER\]";

            // Set up FindReplaceOptions with a custom callback.
            FindReplaceOptions options = new FindReplaceOptions(new InsertNodeCallback());

            // Perform the find-and-replace operation.
            doc.Range.Replace(new Regex(placeholder), string.Empty, options);

            // Save the modified document.
            doc.Save("Output.docx");
        }

        // Callback that replaces the matched text with a new node.
        private class InsertNodeCallback : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                // Create the node to insert (a new paragraph with some text).
                Paragraph newParagraph = new Paragraph(args.MatchNode.Document);
                Run run = new Run(args.MatchNode.Document, "This is the inserted paragraph.");
                newParagraph.AppendChild(run);

                // Import the new node into the destination document.
                NodeImporter importer = new NodeImporter(
                    newParagraph.Document,
                    args.MatchNode.Document,
                    ImportFormatMode.KeepSourceFormatting);
                Node importedNode = importer.ImportNode(newParagraph, true);

                // Insert the imported node after the paragraph that contains the match.
                Paragraph matchParagraph = (Paragraph)args.MatchNode.ParentNode;
                CompositeNode parent = matchParagraph.ParentNode;
                parent.InsertAfter(importedNode, matchParagraph);

                // Remove the original placeholder text.
                // Since we passed an empty replacement string, the placeholder text is already removed.
                // Return Skip to prevent the default replacement behavior.
                return ReplaceAction.Skip;
            }
        }
    }
}
