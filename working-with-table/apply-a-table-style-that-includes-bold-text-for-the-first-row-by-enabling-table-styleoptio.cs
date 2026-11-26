using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // First row (header) – will be formatted as the first row of the style.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in table style.
        table.Style = doc.Styles["Table Grid"];

        // Enable the FirstRow option so the style treats the first row specially
        // (e.g., bold if the style defines it). The enum member is FirstRow, not FirstRowBold.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document to disk.
        string outputPath = "TableStyleFirstRowBold.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
