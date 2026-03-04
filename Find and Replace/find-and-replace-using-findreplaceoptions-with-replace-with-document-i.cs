using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the main document that contains a placeholder like [MY_DOCUMENT].
        Document mainDoc = new Document("Input.docx");

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentCallback();

        // Replace the placeholder with the contents of another document.
        // The placeholder is identified by a regular expression.
        mainDoc.Range.Replace(new Regex(@"\[MY_DOCUMENT\]"), string.Empty, options);

        // Save the modified document.
        mainDoc.Save("Output.docx");
    }

    // Callback that inserts another document at the location of the match.
    private class InsertDocumentCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document insertDoc = new Document("Insert.docx");

            // The match node's parent is a Paragraph; we will insert after it.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the whole document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, insertDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes from docToInsert after insertionDestination.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // Ensure the destination is a paragraph or a table.
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            // Import nodes from the source document into the destination document.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}
