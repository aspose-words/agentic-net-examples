using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceWithDocument
{
    class Program
    {
        static void Main()
        {
            // Load the document that contains the placeholder.
            Document mainDoc = new Document("Input.docx");

            // Set up FindReplaceOptions with a custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertDocumentAtReplaceHandler()
            };

            // Replace the placeholder "[MY_DOCUMENT]" with the contents of another document.
            mainDoc.Range.Replace(new Regex(@"\[MY_DOCUMENT\]"), string.Empty, options);

            // Save the modified document.
            mainDoc.Save("Output.docx");
        }
    }

    // Callback that inserts another document at the location of the match.
    class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document("Insert.docx");

            // The match is inside a Run; its parent is a Paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after insertionDestination (Paragraph or Table).
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

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
