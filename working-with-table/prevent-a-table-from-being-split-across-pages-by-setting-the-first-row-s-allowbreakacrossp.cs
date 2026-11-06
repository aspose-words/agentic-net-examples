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

        // Build a simple 3‑column table with two rows.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.InsertCell();
        builder.Writeln("Data 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first row of the table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.FirstRow;

        // Prevent the first row from being split across pages.
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document.
        string outputPath = "TableNoBreak.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
