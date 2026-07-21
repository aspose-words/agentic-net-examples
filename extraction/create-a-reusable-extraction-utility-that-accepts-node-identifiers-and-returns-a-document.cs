using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionUtility
{
    // Extracts content based on node type and identifier.
    // nodeType: "Paragraph", "Bookmark", "Table"
    // identifier: for Paragraph/Table - zero‑based index as string; for Bookmark - bookmark name.
    public static Document ExtractContent(Document source, string nodeType, string identifier)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrEmpty(nodeType)) throw new ArgumentException("Node type is required.", nameof(nodeType));
        if (string.IsNullOrEmpty(identifier)) throw new ArgumentException("Identifier is required.", nameof(identifier));

        // Create an empty destination document.
        Document result = new Document();
        result.RemoveAllChildren();

        // Add a section and body to the destination document.
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Use NodeImporter for efficient node import with style handling.
        NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

        switch (nodeType.Trim().ToLowerInvariant())
        {
            case "paragraph":
                if (!int.TryParse(identifier, out int paraIndex))
                    throw new ArgumentException("Paragraph identifier must be an integer index.", nameof(identifier));

                Paragraph sourcePara = source.FirstSection.Body.Paragraphs[paraIndex];
                if (sourcePara == null)
                    throw new InvalidOperationException($"Paragraph at index {paraIndex} not found.");

                // Import the paragraph into the destination document.
                Node importedPara = importer.ImportNode(sourcePara, true);
                body.AppendChild(importedPara);
                break;

            case "bookmark":
                Bookmark bookmark = source.Range.Bookmarks[identifier];
                if (bookmark == null)
                    throw new InvalidOperationException($"Bookmark \"{identifier}\" not found.");

                // Build a new paragraph in the destination document.
                Paragraph para = new Paragraph(result);

                // Traverse nodes between BookmarkStart and BookmarkEnd.
                Node currentNode = bookmark.BookmarkStart.NextSibling;
                while (currentNode != null && currentNode != bookmark.BookmarkEnd)
                {
                    Node importedNode = importer.ImportNode(currentNode, true);
                    para.AppendChild(importedNode);
                    currentNode = currentNode.NextSibling;
                }

                body.AppendChild(para);
                break;

            case "table":
                if (!int.TryParse(identifier, out int tableIndex))
                    throw new ArgumentException("Table identifier must be an integer index.", nameof(identifier));

                NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
                if (tableIndex < 0 || tableIndex >= tables.Count)
                    throw new InvalidOperationException($"Table at index {tableIndex} not found.");

                Table sourceTable = tables[tableIndex] as Table;
                if (sourceTable == null)
                    throw new InvalidOperationException("Failed to cast node to Table.");

                // Import the table into the destination document.
                Node importedTable = importer.ImportNode(sourceTable, true);
                body.AppendChild(importedTable);
                break;

            default:
                throw new NotSupportedException($"Node type \"{nodeType}\" is not supported.");
        }

        return result;
    }

    // Helper to create a sample document containing paragraphs, a table, and bookmarks.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraphs
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Table
        builder.StartTable();
        builder.InsertCell(); builder.Write("A1");
        builder.InsertCell(); builder.Write("B1");
        builder.EndRow();
        builder.InsertCell(); builder.Write("A2");
        builder.InsertCell(); builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Bookmarks
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("This is text inside the bookmark.");
        builder.EndBookmark("SampleBookmark");

        doc.Save(filePath);
    }

    public static void Main()
    {
        // Create a sample document.
        string samplePath = "sample.docx";
        CreateSampleDocument(samplePath);

        Document sourceDoc = new Document(samplePath);

        // Extract the second paragraph (index 1).
        Document paraDoc = ExtractContent(sourceDoc, "Paragraph", "1");
        string paraOutput = "extracted-paragraph.docx";
        paraDoc.Save(paraOutput);
        if (!File.Exists(paraOutput))
            throw new InvalidOperationException("Paragraph extraction failed.");

        // Extract the bookmark content.
        Document bookmarkDoc = ExtractContent(sourceDoc, "Bookmark", "SampleBookmark");
        string bookmarkOutput = "extracted-bookmark.docx";
        bookmarkDoc.Save(bookmarkOutput);
        if (!File.Exists(bookmarkOutput))
            throw new InvalidOperationException("Bookmark extraction failed.");

        // Extract the first table (index 0).
        Document tableDoc = ExtractContent(sourceDoc, "Table", "0");
        string tableOutput = "extracted-table.docx";
        tableDoc.Save(tableOutput);
        if (!File.Exists(tableOutput))
            throw new InvalidOperationException("Table extraction failed.");

        // All extractions completed successfully.
    }
}
