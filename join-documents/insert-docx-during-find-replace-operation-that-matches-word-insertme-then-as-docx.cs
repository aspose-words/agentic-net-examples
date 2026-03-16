using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertDocOnFindReplace
{
    static void Main()
    {
        // Load the document that contains the placeholder "INSERTME".
        Document mainDoc = new Document("MainDocument.docx");

        // Configure find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentHandler();

        // Replace the placeholder. The replacement string is empty because the callback
        // will insert the new document and remove the placeholder paragraph.
        mainDoc.Range.Replace(new Regex("\\bINSERTME\\b"), "", options);

        // Save the result.
        mainDoc.Save("Result.docx");
    }

    // Callback that inserts another document at each match.
    private class InsertDocumentHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document("DocumentToInsert.docx");

            // The match node is a Run; its parent is the Paragraph that contains the placeholder.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replace action because we have already handled the insertion.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after insertionDestination (paragraph or table).
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph &&
                insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Destination must be a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(
                docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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
