using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionUtility
{
    private readonly Document _source;

    public ExtractionUtility(Document source)
    {
        _source = source ?? throw new ArgumentNullException(nameof(source));
    }

    // Extracts the content of a bookmark into a new document.
    public Document ExtractBookmark(string bookmarkName)
    {
        if (string.IsNullOrEmpty(bookmarkName))
            throw new ArgumentException("Bookmark name must be provided.", nameof(bookmarkName));

        Bookmark bookmark = _source.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // Create a new empty document with a proper structure.
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Preserve the bookmark text as a paragraph.
        Paragraph para = new Paragraph(result);
        para.AppendChild(new Run(result, bookmark.Text));
        body.AppendChild(para);

        return result;
    }

    // Extracts a range of paragraphs (inclusive) by their zero‑based indices.
    public Document ExtractParagraphRange(int startIndex, int endIndex)
    {
        if (startIndex < 0 || endIndex < startIndex)
            throw new ArgumentException("Invalid paragraph range.");

        ParagraphCollection sourceParas = _source.FirstSection.Body.Paragraphs;
        if (endIndex >= sourceParas.Count)
            throw new ArgumentOutOfRangeException(nameof(endIndex), "End index exceeds paragraph count.");

        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        for (int i = startIndex; i <= endIndex; i++)
        {
            // Import the paragraph into the destination document before appending.
            Node imported = result.ImportNode(sourceParas[i], true);
            body.AppendChild(imported);
        }

        return result;
    }

    // Extracts a table by its zero‑based index.
    public Document ExtractTable(int tableIndex)
    {
        if (tableIndex < 0)
            throw new ArgumentException("Table index must be non‑negative.", nameof(tableIndex));

        NodeCollection tables = _source.GetChildNodes(NodeType.Table, true);
        if (tableIndex >= tables.Count)
            throw new InvalidOperationException($"Table at index {tableIndex} not found.");

        Table table = tables[tableIndex] as Table;
        if (table == null)
            throw new InvalidOperationException($"Node at index {tableIndex} is not a table.");

        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Import the table into the new document.
        Node imported = result.ImportNode(table, true);
        body.AppendChild(imported);

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with paragraphs, a table and bookmarks.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Paragraphs
        builder.Writeln("Paragraph 0 – introduction.");
        builder.Writeln("Paragraph 1 – first content.");
        builder.Writeln("Paragraph 2 – second content.");
        builder.Writeln("Paragraph 3 – conclusion.");

        // Insert a table
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.InsertCell();
        builder.Write("Cell B2");
        builder.EndTable();

        // Bookmarks around the second paragraph
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("This text is inside the bookmark.");
        builder.EndBookmark("SampleBookmark");

        // Save the source document.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and use the extraction utility.
        // -----------------------------------------------------------------
        Document loaded = new Document(sourcePath);
        ExtractionUtility extractor = new ExtractionUtility(loaded);

        // Extract bookmark content.
        Document bookmarkDoc = extractor.ExtractBookmark("SampleBookmark");
        const string bookmarkPath = "bookmark_extracted.docx";
        bookmarkDoc.Save(bookmarkPath);
        if (!File.Exists(bookmarkPath))
            throw new InvalidOperationException("Bookmark extraction failed – output file not created.");

        // Extract paragraphs 1 through 2 (inclusive).
        Document paraRangeDoc = extractor.ExtractParagraphRange(1, 2);
        const string paraRangePath = "paragraph_range_extracted.docx";
        paraRangeDoc.Save(paraRangePath);
        if (!File.Exists(paraRangePath))
            throw new InvalidOperationException("Paragraph range extraction failed – output file not created.");

        // Extract the first table.
        Document tableDoc = extractor.ExtractTable(0);
        const string tablePath = "table_extracted.docx";
        tableDoc.Save(tablePath);
        if (!File.Exists(tablePath))
            throw new InvalidOperationException("Table extraction failed – output file not created.");

        // -----------------------------------------------------------------
        // 3. Serialize simple metadata about the extraction to JSON (optional).
        // -----------------------------------------------------------------
        var metadata = new
        {
            BookmarkExtractionFile = bookmarkPath,
            ParagraphRangeExtractionFile = paraRangePath,
            TableExtractionFile = tablePath,
            ExtractionTime = DateTime.UtcNow
        };

        string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);
        const string jsonPath = "extraction_metadata.json";
        File.WriteAllText(jsonPath, json);
        if (!File.Exists(jsonPath))
            throw new InvalidOperationException("Failed to write extraction metadata JSON.");

        // Indicate successful completion.
        Console.WriteLine("Extraction completed successfully.");
    }
}
