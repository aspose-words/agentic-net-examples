using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionUtility
{
    // Extracts content based on node type and identifier.
    // nodeType: "Paragraph", "Bookmark", or "Table"
    // identifier: paragraph index (0‑based) or bookmark name or table index (0‑based)
    public static Document ExtractContent(Document source, string nodeType, string identifier)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrEmpty(nodeType)) throw new ArgumentException("Node type is required.", nameof(nodeType));
        if (string.IsNullOrEmpty(identifier)) throw new ArgumentException("Identifier is required.", nameof(identifier));

        // Create a clean destination document.
        Document result = new Document();
        result.RemoveAllChildren();
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        switch (nodeType.Trim().ToLowerInvariant())
        {
            case "paragraph":
                {
                    if (!int.TryParse(identifier, out int paraIndex))
                        throw new ArgumentException("Paragraph identifier must be an integer index.", nameof(identifier));

                    Paragraph paragraph = source.FirstSection.Body.Paragraphs[paraIndex];
                    if (paragraph == null)
                        throw new InvalidOperationException($"Paragraph at index {paraIndex} not found.");

                    // Import the paragraph into the destination document before appending.
                    Node imported = result.ImportNode(paragraph, true);
                    body.AppendChild(imported);
                    break;
                }
            case "bookmark":
                {
                    Bookmark bookmark = source.Range.Bookmarks[identifier];
                    if (bookmark == null)
                        throw new InvalidOperationException($"Bookmark \"{identifier}\" not found.");

                    // Create a new paragraph containing the bookmark text.
                    Paragraph para = new Paragraph(result);
                    Run run = new Run(result, bookmark.Text);
                    para.AppendChild(run);
                    body.AppendChild(para);
                    break;
                }
            case "table":
                {
                    if (!int.TryParse(identifier, out int tableIndex))
                        throw new ArgumentException("Table identifier must be an integer index.", nameof(identifier));

                    NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
                    if (tableIndex < 0 || tableIndex >= tables.Count)
                        throw new InvalidOperationException($"Table at index {tableIndex} not found.");

                    Table table = tables[tableIndex] as Table;
                    if (table == null)
                        throw new InvalidOperationException($"Table at index {tableIndex} not found.");

                    // Import the table into the destination document before appending.
                    Node imported = result.ImportNode(table, true);
                    body.AppendChild(imported);
                    break;
                }
            default:
                throw new ArgumentException($"Unsupported node type \"{nodeType}\".", nameof(nodeType));
        }

        return result;
    }
}

public class Program
{
    public static void Main()
    {
        // Build a sample source document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Bookmark surrounding a paragraph.
        builder.StartBookmark("SampleBookmark");
        builder.Writeln("Paragraph inside bookmark.");
        builder.EndBookmark("SampleBookmark");

        // Table.
        builder.StartTable();
        builder.InsertCell(); builder.Write("A1");
        builder.InsertCell(); builder.Write("B1");
        builder.EndRow();
        builder.InsertCell(); builder.Write("A2");
        builder.InsertCell(); builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // Load the document for extraction.
        Document loaded = new Document(sourcePath);

        // Define extraction requests.
        var requests = new List<(string NodeType, string Identifier, string OutputFile)>
        {
            ("Paragraph", "1", "extracted-paragraph.docx"),
            ("Bookmark", "SampleBookmark", "extracted-bookmark.docx"),
            ("Table", "0", "extracted-table.docx")
        };

        // Collect report data.
        var report = new List<object>();

        foreach (var (nodeType, identifier, outputFile) in requests)
        {
            Document extracted = ExtractionUtility.ExtractContent(loaded, nodeType, identifier);
            extracted.Save(outputFile);

            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Failed to create output file \"{outputFile}\".");

            report.Add(new { NodeType = nodeType, Identifier = identifier, OutputFile = outputFile });
        }

        // Serialize a simple JSON report of the extraction operations.
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string reportPath = "extraction-report.json";
        File.WriteAllText(reportPath, jsonReport);

        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Extraction report was not created.");
    }
}
