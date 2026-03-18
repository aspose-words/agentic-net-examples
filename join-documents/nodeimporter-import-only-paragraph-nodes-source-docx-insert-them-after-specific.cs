using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ImportParagraphsAfterBookmark
{
    static void Main()
    {
        // Load the destination document that contains the target bookmark.
        Document dstDoc = new Document("Destination.docx");

        // Load the source document whose paragraphs we want to import.
        Document srcDoc = new Document("Source.docx");

        // Locate the bookmark by name. Replace "MyBookmark" with the actual bookmark name.
        Bookmark bookmark = dstDoc.Range.Bookmarks["MyBookmark"];
        if (bookmark == null)
            throw new ArgumentException("Bookmark 'MyBookmark' not found in the destination document.");

        // The node after which we will insert the imported paragraphs.
        // BookmarkStart.ParentNode is usually a Paragraph.
        Node insertionDestination = bookmark.BookmarkStart.ParentNode;

        // Ensure the insertion point is a paragraph or a table (required by the importer logic).
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The bookmark must be placed inside a paragraph or a table.");

        // The parent story (body) that will receive the new nodes.
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Create a NodeImporter to efficiently import nodes while preserving formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Iterate through all sections of the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            // Iterate through each node in the section's body.
            foreach (Node srcNode in srcSection.Body)
            {
                // Process only paragraph nodes.
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph srcParagraph = (Paragraph)srcNode;

                // Skip the final empty paragraph of a section (it would create an extra blank line).
                if (srcParagraph.IsEndOfSection && !srcParagraph.HasChildNodes)
                    continue;

                // Import the paragraph (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcParagraph, true);

                // Insert the imported paragraph after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);

                // Update the insertion point so the next paragraph is placed after this one.
                insertionDestination = importedNode;
            }
        }

        // Save the modified destination document.
        dstDoc.Save("Result.docx");
    }
}
