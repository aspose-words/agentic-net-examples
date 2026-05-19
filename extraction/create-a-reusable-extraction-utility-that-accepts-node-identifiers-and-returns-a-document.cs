using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionUtility
{
    // Extracts the content that lies inside a bookmark and returns it as a new document.
    public static Document ExtractByBookmark(Document source, string bookmarkName)
    {
        Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        Node startNode = bookmark.BookmarkStart;
        Node endNode = bookmark.BookmarkEnd;

        // Collect all nodes that are between the start and end markers.
        List<Node> nodesInRange = new List<Node>();
        Node cur = startNode.NextSibling;
        while (cur != null && cur != endNode)
        {
            nodesInRange.Add(cur);
            cur = cur.NextSibling;
        }

        if (nodesInRange.Count == 0)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' contains no nodes.");

        // Build the result document with a clean structure.
        Document result = new Document();
        result.RemoveAllChildren();
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        Paragraph? inlineParagraph = null; // Holds runs that belong to the same paragraph.

        foreach (Node node in nodesInRange)
        {
            // Import the node into the destination document.
            Node imported = result.ImportNode(node, true);

            // Block level nodes can be appended directly to the body.
            if (imported.NodeType == NodeType.Paragraph || imported.NodeType == NodeType.Table)
            {
                // Flush any pending inline paragraph first.
                if (inlineParagraph != null)
                {
                    body.AppendChild(inlineParagraph);
                    inlineParagraph = null;
                }

                body.AppendChild(imported);
            }
            else // Inline nodes (Run, Field, etc.) must be placed inside a paragraph.
            {
                if (inlineParagraph == null)
                    inlineParagraph = new Paragraph(result);

                inlineParagraph.AppendChild(imported);
            }
        }

        // Append any remaining inline paragraph.
        if (inlineParagraph != null)
            body.AppendChild(inlineParagraph);

        return result;
    }

    // Extracts a range of paragraphs (inclusive) by their zero‑based indices.
    public static Document ExtractParagraphRange(Document source, int startIndex, int endIndex)
    {
        ParagraphCollection paragraphs = source.FirstSection.Body.Paragraphs;
        if (startIndex < 0 || endIndex >= paragraphs.Count || startIndex > endIndex)
            throw new ArgumentOutOfRangeException("Invalid paragraph range.");

        Document result = new Document();
        result.RemoveAllChildren();
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = paragraphs[i];
            Paragraph cloned = (Paragraph)para.Clone(true);
            Node imported = result.ImportNode(cloned, true);
            body.AppendChild(imported);
        }

        return result;
    }

    // Extracts a table by its zero‑based index in the document.
    public static Document ExtractTable(Document source, int tableIndex)
    {
        NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentOutOfRangeException("Table index out of range.");

        Table table = tables[tableIndex] as Table;
        if (table == null)
            throw new InvalidOperationException("Selected node is not a table.");

        Document result = new Document();
        result.RemoveAllChildren();
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        Table clonedTable = (Table)table.Clone(true);
        Node imported = result.ImportNode(clonedTable, true);
        body.AppendChild(imported);

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs, a bookmark and a table.
        Document sample = new Document();
        DocumentBuilder builder = new DocumentBuilder(sample);

        builder.Writeln("Paragraph 1");
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("Paragraph inside bookmark");
        builder.EndBookmark("SampleBookmark");
        builder.Writeln("Paragraph 3");

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell(); builder.Write("A1");
        builder.InsertCell(); builder.Write("B1");
        builder.EndRow();
        builder.InsertCell(); builder.Write("A2");
        builder.InsertCell(); builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document.
        const string sourcePath = "sample.docx";
        sample.Save(sourcePath);

        // Load the document for extraction.
        Document loaded = new Document(sourcePath);

        // 1. Extract content inside the bookmark.
        Document bookmarkExtract = ExtractionUtility.ExtractByBookmark(loaded, "SampleBookmark");
        const string bookmarkPath = "extracted-bookmark.docx";
        bookmarkExtract.Save(bookmarkPath);
        if (!File.Exists(bookmarkPath))
            throw new InvalidOperationException("Bookmark extraction failed.");

        // 2. Extract paragraphs 0 through 2 (the three paragraphs).
        Document paragraphExtract = ExtractionUtility.ExtractParagraphRange(loaded, 0, 2);
        const string paragraphPath = "extracted-paragraphs.docx";
        paragraphExtract.Save(paragraphPath);
        if (!File.Exists(paragraphPath))
            throw new InvalidOperationException("Paragraph range extraction failed.");

        // 3. Extract the first table in the document.
        Document tableExtract = ExtractionUtility.ExtractTable(loaded, 0);
        const string tablePath = "extracted-table.docx";
        tableExtract.Save(tablePath);
        if (!File.Exists(tablePath))
            throw new InvalidOperationException("Table extraction failed.");

        // Simple console output to indicate success (no interactive input required).
        Console.WriteLine("Extraction completed successfully.");
    }
}
