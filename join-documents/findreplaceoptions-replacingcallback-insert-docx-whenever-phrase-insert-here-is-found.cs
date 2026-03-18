using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the destination document that contains the placeholder "INSERT_HERE".
        Document mainDoc = new Document("Destination.docx");

        // Configure find/replace options with a custom callback that will insert another document.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocHandler();

        // Replace the placeholder with an empty string; the callback will handle the insertion.
        mainDoc.Range.Replace(new Regex("INSERT_HERE"), "", options);

        // Save the modified document.
        mainDoc.Save("Result.docx");
    }

    // Callback that inserts a separate DOCX at the location of each match.
    private class InsertDocHandler : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document("Insert.docx");

            // The match node is inside a Run; its parent paragraph is the insertion point.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;

            // Insert the entire sub‑document after the paragraph containing the match.
            InsertDocument(para, subDoc);

            // Remove the paragraph that held the placeholder text.
            para.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }

        // Helper method that inserts all nodes of another document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Destination must be a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;
            // NodeImporter now lives directly under Aspose.Words namespace.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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

                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}
