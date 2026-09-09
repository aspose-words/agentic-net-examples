using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionUtility
{
    // Creates an empty document with a single section and body.
    private static Document CreateEmptyDocument()
    {
        Document doc = new Document();
        doc.RemoveAllChildren();

        Section section = new Section(doc);
        doc.AppendChild(section);

        Body body = new Body(doc);
        section.AppendChild(body);

        return doc;
    }

    // Extracts content based on a simple identifier format.
    // Supported formats:
    //   Paragraph:{zero‑based index}
    //   Table:{zero‑based index}
    //   Bookmark:{bookmark name}
    public static Document Extract(string identifier, Document source)
    {
        if (identifier == null) throw new ArgumentNullException(nameof(identifier));
        if (source == null) throw new ArgumentNullException(nameof(source));

        Document result = CreateEmptyDocument();

        if (identifier.StartsWith("Paragraph:", StringComparison.OrdinalIgnoreCase))
        {
            string[] parts = identifier.Split(':');
            if (parts.Length != 2 || !int.TryParse(parts[1], out int index))
                throw new ArgumentException("Invalid paragraph identifier.", nameof(identifier));

            Paragraph paragraph = source.FirstSection?.Body?.Paragraphs[index];
            if (paragraph == null)
                throw new InvalidOperationException($"Paragraph at index {index} not found.");

            // Import the paragraph into the result document.
            Node imported = result.ImportNode(paragraph, true);
            result.FirstSection.Body.AppendChild(imported);
        }
        else if (identifier.StartsWith("Table:", StringComparison.OrdinalIgnoreCase))
        {
            string[] parts = identifier.Split(':');
            if (parts.Length != 2 || !int.TryParse(parts[1], out int index))
                throw new ArgumentException("Invalid table identifier.", nameof(identifier));

            NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
            if (index < 0 || index >= tables.Count)
                throw new InvalidOperationException($"Table at index {index} not found.");

            Table table = tables[index] as Table;
            if (table == null)
                throw new InvalidOperationException($"Node at index {index} is not a table.");

            // Import the table into the result document.
            Node imported = result.ImportNode(table, true);
            result.FirstSection.Body.AppendChild(imported);
        }
        else if (identifier.StartsWith("Bookmark:", StringComparison.OrdinalIgnoreCase))
        {
            string[] parts = identifier.Split(new[] { ':' }, 2);
            if (parts.Length != 2)
                throw new ArgumentException("Invalid bookmark identifier.", nameof(identifier));

            string bookmarkName = parts[1];
            Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                throw new InvalidOperationException($"Bookmark \"{bookmarkName}\" not found.");

            // Create a paragraph containing the bookmark text.
            Paragraph para = new Paragraph(result);
            para.AppendChild(new Run(result, bookmark.Text));
            result.FirstSection.Body.AppendChild(para);
        }
        else
        {
            throw new ArgumentException("Unsupported identifier type.", nameof(identifier));
        }

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Add paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Add a table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Add a bookmark around some text.
        builder.StartBookmark("SampleBookmark");
        builder.Write("Text inside bookmark.");
        builder.EndBookmark("SampleBookmark");

        // Save the source document (demonstrates loading from file).
        string sourcePath = "source.docx";
        source.Save(sourcePath);

        // Load the document (demonstrates loading from file).
        Document loaded = new Document(sourcePath);

        // Extract a paragraph (index 1 -> second paragraph).
        Document paragraphDoc = ExtractionUtility.Extract("Paragraph:1", loaded);
        string paragraphPath = "extracted-paragraph.docx";
        paragraphDoc.Save(paragraphPath);
        if (!File.Exists(paragraphPath))
            throw new InvalidOperationException("Paragraph extraction failed.");

        // Extract the first table (index 0).
        Document tableDoc = ExtractionUtility.Extract("Table:0", loaded);
        string tablePath = "extracted-table.docx";
        tableDoc.Save(tablePath);
        if (!File.Exists(tablePath))
            throw new InvalidOperationException("Table extraction failed.");

        // Extract the bookmark content.
        Document bookmarkDoc = ExtractionUtility.Extract("Bookmark:SampleBookmark", loaded);
        string bookmarkPath = "extracted-bookmark.docx";
        bookmarkDoc.Save(bookmarkPath);
        if (!File.Exists(bookmarkPath))
            throw new InvalidOperationException("Bookmark extraction failed.");

        // All extractions completed successfully.
    }
}
