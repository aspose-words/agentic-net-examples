using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the main document that contains the placeholder tag.
        Document mainDoc = new Document("Document insertion destination.docx");

        // Configure find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentAtReplaceHandler();

        // Replace the placeholder tag (e.g. [MY_DOCUMENT]) and insert the sub‑document.
        mainDoc.Range.Replace(new Regex(@"\[MY_DOCUMENT\]"), string.Empty, options);

        // Save the resulting document as DOCX.
        mainDoc.Save("InsertDocumentResult.docx");
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document("Document.docx");

            // The match is inside a paragraph; get that paragraph.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the paragraph containing the match.
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip the default replace action because we have already handled the insertion.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after insertionDestination (paragraph or table).
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph &&
                insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be either a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            // Use NodeImporter for efficient import and style handling.
            NodeImporter importer = new NodeImporter(
                docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            // Iterate through all block‑level nodes of the source document.
            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph p = (Paragraph)srcNode;
                        if (p.IsEndOfSection && !p.HasChildNodes)
                            continue;
                    }

                    // Import the node into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert after the current insertion point and update the pointer.
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}
