using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ImportParagraphsAfterBookmark
{
    static void Main()
    {
        // Create a destination document with a bookmark named "MyBookmark".
        Document dstDoc = new Document();
        Section dstSection = new Section(dstDoc);
        dstDoc.AppendChild(dstSection);
        Body dstBody = new Body(dstDoc);
        dstSection.AppendChild(dstBody);

        Paragraph dstParagraph = new Paragraph(dstDoc);
        dstBody.AppendChild(dstParagraph);
        // Insert a bookmark start.
        BookmarkStart bookmarkStart = new BookmarkStart(dstDoc, "MyBookmark");
        dstParagraph.AppendChild(bookmarkStart);
        // Add some text inside the bookmark.
        Run run = new Run(dstDoc, "This is the bookmark location.");
        dstParagraph.AppendChild(run);
        // Insert a bookmark end.
        BookmarkEnd bookmarkEnd = new BookmarkEnd(dstDoc, "MyBookmark");
        dstParagraph.AppendChild(bookmarkEnd);

        // Create a source document with a few paragraphs.
        Document srcDoc = new Document();
        Section srcSection = new Section(srcDoc);
        srcDoc.AppendChild(srcSection);
        Body srcBody = new Body(srcDoc);
        srcSection.AppendChild(srcBody);

        for (int i = 1; i <= 3; i++)
        {
            Paragraph p = new Paragraph(srcDoc);
            Run r = new Run(srcDoc, $"Source paragraph {i}");
            p.AppendChild(r);
            srcBody.AppendChild(p);
        }

        // Locate the bookmark in the destination document.
        Bookmark bookmark = dstDoc.Range.Bookmarks["MyBookmark"];
        if (bookmark == null)
            throw new ArgumentException("Bookmark 'MyBookmark' not found in the destination document.");

        // The node after which we will insert the imported paragraphs.
        Node insertionDestination = bookmark.BookmarkStart.ParentNode;

        // Ensure the insertion point is a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The bookmark must be placed inside a paragraph or a table.");

        // The parent story (body) that will receive the new nodes.
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Create a NodeImporter to import nodes while preserving formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Iterate through all sections of the source document.
        foreach (Section srcSectionIter in srcDoc.Sections)
        {
            // Iterate through each node in the section's body.
            foreach (Node srcNode in srcSectionIter.Body)
            {
                if (srcNode.NodeType != NodeType.Paragraph)
                    continue;

                Paragraph srcParagraph = (Paragraph)srcNode;

                // Skip the final empty paragraph of a section.
                if (srcParagraph.IsEndOfSection && !srcParagraph.HasChildNodes)
                    continue;

                // Import the paragraph into the destination document.
                Node importedNode = importer.ImportNode(srcParagraph, true);

                // Insert the imported paragraph after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);

                // Update the insertion point for the next insertion.
                insertionDestination = importedNode;
            }
        }

        // Save the resulting document.
        dstDoc.Save("Result.docx");
        Console.WriteLine("Result.docx has been created successfully.");
    }
}
