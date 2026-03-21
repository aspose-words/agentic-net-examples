using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class BookmarkExtractor
{
    static void Main()
    {
        // Create a source document with a bookmark for demonstration purposes.
        Document srcDoc = new Document();
        const string bookmarkName = "ContentBookmark";

        // Build a paragraph: "Before [bookmark start]Extracted content.[bookmark end] After"
        Paragraph para = new Paragraph(srcDoc);
        para.AppendChild(new Run(srcDoc, "Before "));
        BookmarkStart bStart = new BookmarkStart(srcDoc, bookmarkName);
        para.AppendChild(bStart);
        para.AppendChild(new Run(srcDoc, "Extracted content."));
        BookmarkEnd bEnd = new BookmarkEnd(srcDoc, bookmarkName);
        para.AppendChild(bEnd);
        para.AppendChild(new Run(srcDoc, " After"));
        srcDoc.FirstSection.Body.AppendChild(para);

        // Retrieve the bookmark.
        Bookmark bookmark = srcDoc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // Nodes that lie between the start and end of the bookmark.
        Node startNode = bookmark.BookmarkStart;
        Node endNode = bookmark.BookmarkEnd;

        // Collect the nodes to extract (excluding the bookmark markers themselves).
        List<Node> nodesToExtract = new List<Node>();
        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
            nodesToExtract.Add(cur);

        // -----------------------------------------------------------------
        // 1. Create a new document that will contain the extracted content.
        // -----------------------------------------------------------------
        Document extractedDoc = new Document();
        CompositeNode dstStory = extractedDoc.FirstSection.Body;

        // Use NodeImporter for efficient import of nodes from the source document.
        NodeImporter importer = new NodeImporter(srcDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Create a paragraph in the destination document to hold the imported nodes.
        Paragraph dstParagraph = new Paragraph(extractedDoc);
        foreach (Node node in nodesToExtract)
        {
            Node imported = importer.ImportNode(node, true);
            dstParagraph.AppendChild(imported);
        }
        dstStory.AppendChild(dstParagraph);

        // Save the extracted fragment (optional, for verification).
        extractedDoc.Save("ExtractedContent.docx");

        // ---------------------------------------------------------------
        // 2. Remove the original nodes from the source document.
        // ---------------------------------------------------------------
        foreach (Node node in nodesToExtract)
            node.Remove();

        // ---------------------------------------------------------------
        // 3. Insert a placeholder paragraph where the original content was.
        // ---------------------------------------------------------------
        // Capture the original paragraph that holds the bookmark.
        Paragraph originalParagraph = (Paragraph)startNode.ParentNode;

        // Remove the bookmark markers.
        startNode.Remove();
        endNode.Remove();

        // Create placeholder paragraph.
        Paragraph placeholder = new Paragraph(srcDoc);
        placeholder.AppendChild(new Run(srcDoc, "[Placeholder]"));

        // Insert placeholder after the original paragraph and then remove the original paragraph.
        CompositeNode paragraphParent = (CompositeNode)originalParagraph.ParentNode;
        paragraphParent.InsertAfter(placeholder, originalParagraph);
        originalParagraph.Remove();

        // Save the modified source document.
        srcDoc.Save("OutputWithPlaceholder.docx");
    }
}
