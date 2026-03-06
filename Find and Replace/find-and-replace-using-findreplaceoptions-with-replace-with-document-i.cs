using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWithDocument
{
    static void Main()
    {
        // Load the main document that contains a placeholder like [MY_DOCUMENT].
        Document mainDoc = new Document(@"C:\Data\MainDocument.docx");

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentCallback();

        // Replace the placeholder with the contents of another document.
        mainDoc.Range.Replace(new Regex(@"\[MY_DOCUMENT\]"), string.Empty, options);

        // Save the modified document.
        mainDoc.Save(@"C:\Output\Result.docx");
    }

    // Callback that inserts a document at the location of the match.
    private class InsertDocumentCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(@"C:\Data\InsertDocument.docx");

            // The match is inside a Run; its parent paragraph will be replaced.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the sub‑document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after a given paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // Ensure the destination is a paragraph.
            if (insertionDestination.NodeType != NodeType.Paragraph)
                throw new ArgumentException("Insertion destination must be a paragraph.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            // Import nodes from the source document, preserving formatting.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the final empty paragraph of a section.
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
